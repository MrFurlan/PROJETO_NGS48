<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AnaliseDeCreditoParametros.aspx.vb" Inherits="NGS.Web.UI.AnaliseDeCreditoParametros" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngAnalisesDeCreditoParametros" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAnalisesDeCreditoParametros" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Parâmetros Análise de Crédito
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovoB" Text="Gravar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimparB" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjudaB" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="painelleft" style="width: 185px;">
                <div class="row">
                    <div class="collbl" style="width: 96px;">
                        Ano:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlAno" runat="server" AutoPostBack="True" EnableTheming="True"
                            OnSelectedIndexChanged="ddlAno_SelectedIndexChanged" Width="70px">
                            <asp:ListItem></asp:ListItem>
                            <asp:ListItem Value="2012"></asp:ListItem>
                            <asp:ListItem>2013</asp:ListItem>
                            <asp:ListItem>2014</asp:ListItem>
                            <asp:ListItem>2015</asp:ListItem>
                            <asp:ListItem>2016</asp:ListItem>
                            <asp:ListItem>2017</asp:ListItem>
                            <asp:ListItem>2018</asp:ListItem>
                            <asp:ListItem>2019</asp:ListItem>
                            <asp:ListItem>2020</asp:ListItem>
                            <asp:ListItem>2021</asp:ListItem>
                            <asp:ListItem>2022</asp:ListItem>
                            <asp:ListItem>2023</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="bordagrid" style="width: 180px;">
                    <asp:GridView ID="GridDefinicoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="GridDefinicoes_SelectedIndexChanged"
                        Width="100%">
                        <EditRowStyle BackColor="#999999" />
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                            <asp:BoundField DataField="DefinicaoAno" HeaderText="Definição" ItemStyle-HorizontalAlign="Center" />
                            <asp:BoundField DataField="DataDefinicao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data"
                                ItemStyle-HorizontalAlign="Center" />
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <div class="painelleft" style="width: 760px;">
                <ajaxToolkit:TabContainer ID="TCItensParametro" runat="server" ActiveTabIndex="1"
                    Width="100%">
                    <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" Style="min-height: 405px;">
                        <HeaderTemplate>
                            Parâmetros da Definição
                        </HeaderTemplate>
                        <ContentTemplate>
                            <div class="row">
                                <div class="collbl">
                                    % Contas Abertas:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPercContasAbertas" runat="server" CssClass="txtDecimal" />
                                </div>
                            </div>
                            <div class="painelleft" style="width: 31.5%; margin-right: 5px;">
                                <div class="subtitulodiv">
                                    Limite de Crédito
                                </div>
                            </div>
                            <div class="painelleft" style="width: 32%;">
                                <div class="subtitulodiv">
                                    % Redutor Risco Cultura
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    % Nível A:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPercLimCredA" runat="server" CssClass="txtDecimal" />
                                </div>
                                <div class="collbl">
                                    Risco Alto:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtRiscoAlto" runat="server" CssClass="txtDecimal" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    % Nível B:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPercLimCredB" runat="server" CssClass="txtDecimal" />
                                </div>
                                <div class="collbl">
                                    Risco Médio:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtRiscoMedio" runat="server" CssClass="txtDecimal" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    % Nível C:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPercLimCredC" runat="server" CssClass="txtDecimal" />
                                </div>
                                <div class="collbl">
                                    Risco Baixo:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtRiscoBaixo" runat="server" CssClass="txtDecimal" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Cotação do Dólar:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCotacaoDolar" runat="server" CssClass="txtDecimal" />
                                </div>
                                <div class="collbl">
                                    Custo Arrend. Ha:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCustoArrendamentoha" runat="server" CssClass="txtDecimal" />
                                </div>
                            </div>
                        </ContentTemplate>
                    </ajaxToolkit:TabPanel>
                    <ajaxToolkit:TabPanel ID="TabCulturas" runat="server" HeaderText="Parâmetros Culturas">
                        <ContentTemplate>
                            <div class="subtitulodiv">
                                Parâmetros Culturas
                            </div>
                            <div class="menu_acoes">
                                <div class="acoes">
                                    <ul>
                                        <li class="iconNovo" runat="server">
                                            <asp:LinkButton ID="lnkGravarCultura" Text="Gravar" runat="server" />
                                        </li>
                                        <li class="iconLimpar" runat="server">
                                            <asp:LinkButton ID="lnkLimparCultura" Text="Limpar" runat="server" />
                                        </li>
                                        <li class="iconAjuda" runat="server">
                                            <asp:LinkButton ID="lnkAjudaCultura" Text="Ajuda" runat="server" />
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div>
                                <div class="painelleft" style="width: 55%;">
                                    <div class="row">
                                        <div class="collbl">
                                            Safra:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlSafra" runat="server" Width="250px" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Cultura:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlCultura" runat="server" Width="250px">
                                            </asp:DropDownList>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Produtividade Ha:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtProdutividadeha" runat="server" CssClass="txtDecimal"></asp:TextBox>
                                        </div>
                                    </div>
                                    <div class="row">
                                        *Valores em Moeda Nacional
                                    </div>
                                </div>
                                <div class="painelleft" style="width: 44%;">
                                    <div class="row">
                                        <div class="collbl">
                                            Preço Saco:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtPrecoSaco" runat="server" CssClass="txtDecimal"></asp:TextBox>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Receita Ha:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtReceitaHa" runat="server" CssClass="txtDecimal"></asp:TextBox>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Custo Total Ha:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtCustoTotalHa" runat="server" CssClass="txtDecimal"></asp:TextBox>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Custo Portifolio Ha:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtCustoPortifolioHa" runat="server" CssClass="txtDecimal"></asp:TextBox>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="bordagrid" style="height: 250px;">
                                <asp:GridView ID="gridCulturas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" OnRowCancelingEdit="gridCulturas_RowCancelingEdit"
                                    OnRowDeleting="gridCulturas_RowDeleting" OnRowEditing="gridCulturas_RowEditing"
                                    OnRowUpdating="gridCulturas_RowUpdating">
                                    <EditRowStyle BackColor="#999999" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:TemplateField ShowHeader="False">
                                            <ItemTemplate>
                                                <asp:Button ID="btnEditar" runat="server" CausesValidation="False" Text=" &gt; "
                                                    OnClick="btnEditar_Click" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="CodigoSafra" HeaderText="Safra" />
                                        <asp:BoundField DataField="CodigoCultura" DataFormatString="{0:N0}" HeaderText="ID Cult." />
                                        <asp:BoundField DataField="NomeCultura" HeaderText="Cultura" />
                                        <asp:BoundField DataField="Produtividade" DataFormatString="{0:N2}" HeaderText="Produtividade Ha" />
                                        <asp:BoundField DataField="PrecoSaco" DataFormatString="{0:N2}" HeaderText="Preço Saco" />
                                        <asp:BoundField DataField="Receita" DataFormatString="{0:N2}" HeaderText="Receita" />
                                        <asp:BoundField DataField="CustoTotalHa" DataFormatString="{0:N2}" HeaderText="Custo Total Ha" />
                                        <asp:BoundField DataField="CustoPortifolioHa" DataFormatString="{0:N2}" HeaderText="Custo Portifolio Ha" />
                                        <asp:TemplateField HeaderText="Portifolio">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="imdPortifolio" runat="server" OnClick="ImdPortifolio_Click">
                                                    <asp:Image ID="imgPortifolio" runat="server" Width="16px" Height="16px" ImageUrl="~/Images/ico-mais.gif"
                                                        data-ToolTip="default" ToolTip="Composição do Portifólio" />
                                                </asp:LinkButton>
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Excluir">
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgExcluirCultura" runat="server" data-ToolTip="default" ImageUrl="~/Images/deletar.gif"
                                                    OnClick="imgExcluirCultura_Click" OnClientClick="return confirm('Deseja realmente excluir Esta Safra/Cultura?');"
                                                    Style="border: 0;" ToolTip="Excluir" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                            <div class="subtitulodiv">
                                Parâmetros Portifólio
                            </div>
                            <div class="menu_acoes">
                                <div class="acoes">
                                    <ul>
                                        <li class="iconLimpar" runat="server">
                                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                                        </li>
                                        <li class="iconRelatorio" runat="server">
                                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                                        </li>
                                        <li class="iconAjuda" runat="server">
                                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div class="row">
                                <asp:HiddenField ID="HDSafraCultura" runat="server" />
                                <div class="collbl">
                                    Safra:
                                </div>
                                <div class="coltxt" style="width: 280px;">
                                    <asp:Label ID="lblSafra" runat="server" Style="font-size: 14px;" />
                                </div>
                                <div class="collbl">
                                    Cultura:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="lblCultura" runat="server" Style="font-size: 14px;" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt" style="width: 90%;">
                                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:LinkButton ID="lnkAdicionarGrupoProduto" Text="Adicionar" runat="server" Style="text-decoration: none;
                                        padding: 4px; position: relative; top: 6px;" CssClass="btn" />
                                </div>
                            </div>
                            <br />
                            <div class="bordagrid">
                                <asp:GridView ID="GridCulturaPortifolio" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" ShowFooter="True" Width="700px">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                            <ItemStyle Width="25px" />
                                        </asp:CommandField>
                                        <asp:BoundField DataField="CodigoGrupoProduto" HeaderText="Código">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" Width="50px" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Grupo">
                                            <ItemTemplate>
                                                <%# Eval("GrupoProduto.Descricao")%>
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle Width="300px" />
                                        </asp:TemplateField>
                                        <asp:TemplateField FooterText="Total:" HeaderText="Quantidade">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtQuantidade" runat="server" AutoPostBack="True" CssClass="Integer"
                                                    OnTextChanged="txtQuantidade_TextChanged" Text='<%# Eval("Quantidade") %>' Width="80px" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                            <FooterStyle HorizontalAlign="Right" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Valor em R$">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtValor" runat="server" AutoPostBack="True" CssClass="txtDecimal"
                                                    OnTextChanged="txtValor_TextChanged" Text='<%# Eval("Valor") %>' Width="120px" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Excluir">
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgExcluirPortifolio" runat="server" data-ToolTip="default"
                                                    ImageUrl="~/Images/deletar.gif" OnClick="imgExcluirPortifolio_Click" OnClientClick="return confirm('Deseja realmente excluir este grupo do Portifolio?');"
                                                    Style="border: 0;" ToolTip="Excluir" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </ContentTemplate>
                    </ajaxToolkit:TabPanel>
                    <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                        <HeaderTemplate>
                            Perguntas da Definição
                        </HeaderTemplate>
                        <ContentTemplate>
                            <div class="subtitulodiv">
                                Perguntas Comportamentais
                            </div>
                            <div class="menu_acoes">
                                <div class="acoes">
                                    <ul>
                                        <li class="iconNovo" runat="server">
                                            <asp:LinkButton ID="lnkGravarPergunta" Text="Gravar" runat="server" />
                                        </li>
                                        <li class="iconLimpar" runat="server">
                                            <asp:LinkButton ID="lnkLimparPergunta" Text="Limpar" runat="server" />
                                        </li>
                                        <li class="iconAjuda" runat="server">
                                            <asp:LinkButton ID="lnkAjudaPergunta" Text="Ajuda" runat="server" />
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt">
                                    <asp:CheckBox ID="chkPerguntas" runat="server" AutoPostBack="True" OnCheckedChanged="chkPerguntas_CheckedChanged"
                                        Text="Usar Perguntas Comportamentais" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl" style="width: 60px;">
                                    Codigo:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtId" runat="server" Width="25px" Enabled="False" />
                                </div>
                                <div class="collbl" style="width: 60px;">
                                    Pergunta:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPergunta" runat="server" Width="280px" />
                                </div>
                                <div class="collbl" style="width: 60px;">
                                    Peso:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPeso" runat="server" CssClass="txtDecimal" Width="30px" />
                                </div>
                                <div class="coltxt">
                                    %
                                </div>
                            </div>
                            <div class="bordagrid" style="height: 320px;">
                                <asp:GridView ID="gridPerguntas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    Font-Bold="False" ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="False" ForeColor="White" Height="26px" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:TemplateField ShowHeader="False">
                                            <ItemTemplate>
                                                <asp:Button ID="btnPergunta" runat="server" CausesValidation="False" Text=" &gt; "
                                                    OnClick="btnPergunta_Click" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="CodigoPergunta" HeaderText="Codigo">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Descricao" HeaderText="Pergunta">
                                            <HeaderStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PercPeso" DataFormatString="{0:N2}" HeaderText="% Peso">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Excluir">
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgExcluir" runat="server" data-ToolTip="default" ImageUrl="~/Images/deletar.gif"
                                                    OnClick="imgExcluir_Click" OnClientClick="return confirm('Deseja realmente excluir a  pergunta?');"
                                                    Style="border: 0;" ToolTip="Excluir" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </ContentTemplate>
                    </ajaxToolkit:TabPanel>
                </ajaxToolkit:TabContainer>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
