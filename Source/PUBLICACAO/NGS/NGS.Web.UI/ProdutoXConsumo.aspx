<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="ProdutoXConsumo.aspx.vb" Inherits="NGS.Web.UI.ProdutoXConsumo" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngProdutoXConsumo" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlProdutoXConsumo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Produto X Consumo
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%" Style="margin-top: 4px;">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Lista
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid" style="height: 650px;">
                            <asp:GridView ID="gridProdutoXConsumo" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridProdutoXConsumo_SelectedIndexChanged">
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
                                    <asp:BoundField DataField="Nome" HeaderText="Nome" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
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
                                        <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                            OnClientClick="if(!confirm('Deseja realmente excluir o item selecionado?')) return false;" />
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
                        <div class="subtitulodiv">
                            Dados para Produção
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupoProdutoProducao" runat="server" Enabled="false" AutoPostBack="True" Width="576px" />
                            </div>
                            <div class="coltxt">
                                <asp:LinkButton ID="lnkBuscaProdutoProducao" runat="server" Enabled="false" Height="20px" Width="20px">
                                    <asp:Image ID="imgSearch" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                                        ToolTip="Consultar Produto para Produção" />
                                </asp:LinkButton>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProdutoProducao" runat="server" Enabled="false" AutoPostBack="True" Width="596px" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 70px;">
                            <asp:GridView ID="gridProducao" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
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
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgRemoverProducao" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                OnClick="imgRemoverProducao_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente remover o Produto?');" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <div class="subtitulodiv">
                            Dados para Consumo
                        </div>
                        <div class="row">
                            <div class="painelleft" style="width: 84%;">
                                <div class="collbl">
                                    Percentual:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPercentual" runat="server" Enabled="false" CssClass="txtDecimal4" data-ToolTip="default" ToolTip="Percentual de Consumo." />
                                </div>
                                <div>
                                    <asp:Button ID="btnPercentual" OnClick="btnPercentual_Click" runat="server" Visible="false" UseSubmitBehavior="False"
                                        Text="&gt;" CssClass="btn" />
                                </div>
                            </div>
                            <div class="painelright" style="width: 15%;">
                                <asp:Label ID="lblPercentual" runat="server" Text="" Width="130px" BackColor="#FFFFC0" Font-Bold="True" ForeColor="Red" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupoProdutoConsumo" runat="server" Enabled="false" AutoPostBack="True" Width="576px" />
                            </div>
                            <div class="coltxt">
                                <asp:LinkButton ID="lnkBuscaProdutoConsumo" runat="server" Enabled="false" Height="20px" Width="20px">
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
                                <asp:DropDownList ID="ddlProdutosConsumo" runat="server" Enabled="false" AutoPostBack="True" Width="300px" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 180px;">
                            <asp:GridView ID="gridConsumo" runat="server" AutoGenerateColumns="False" OnSelectedIndexChanged="gridConsumo_SelectedIndexChanged" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
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
                                    <asp:BoundField DataField="Percentual" HeaderText="Percentual" HtmlEncode="False">
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
                            Dados para Insumos
                        </div>
                        <div class="row">
                            <div class="painelleft" style="width: 85%;">
                                <div class="collbl">
                                    Base:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtBase" runat="server" Enabled="false" CssClass="txtNumerico9" data-ToolTip="default" ToolTip="Base do Insumo." />
                                </div>
                                <div>
                                    <asp:Button ID="btnBase" OnClick="btnBase_Click" runat="server" Visible="false" UseSubmitBehavior="False"
                                        Text="&gt;" CssClass="btn" />
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupoProdutoInsumo" runat="server" Enabled="false" AutoPostBack="True" Width="576px" />
                            </div>
                            <div class="coltxt">
                                <asp:LinkButton ID="lnkBuscaProdutoInsumo" runat="server" Enabled="false" Height="20px" Width="20px">
                                    <asp:Image ID="Image2" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                                        ToolTip="Consultar Insumo" />
                                </asp:LinkButton>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProdutosInsumo" runat="server" Enabled="false" AutoPostBack="True" Width="596px" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 100px;">
                            <asp:GridView ID="gridInsumo" runat="server" AutoGenerateColumns="False" OnSelectedIndexChanged="gridInsumo_SelectedIndexChanged" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
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
                                    <asp:BoundField DataField="Base" HeaderText="Base" HtmlEncode="False">
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
    </asp:UpdatePanel>
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>

