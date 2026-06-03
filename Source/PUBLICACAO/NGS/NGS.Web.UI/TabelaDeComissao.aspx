<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="TabelaDeComissao.aspx.vb" Inherits="NGS.Web.UI.TabelaDeComissao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngTabelaDeComissao" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <div class="titulodiv">
        Tabela de Comissão
    </div>
    <asp:UpdatePanel ID="updpnlTabelaDeComissao" runat="server">
        <ContentTemplate>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Tabela de Comissão
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                                    </li>
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
                            <div class="collbl">
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigo" runat="server" data-ToolTip="default" ToolTip="Número do registro do produto." />
                                <asp:HiddenField ID="txtCodigoTabela" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricao" runat="server" Width="500px" data-ToolTip="default" ToolTip="Nome do produto." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Safra:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSafra" runat="server" Width="508px" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridTabelaDeComissao" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridTabelaDeComissao_SelectedIndexChanged">
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
                                    <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                        <ItemStyle HorizontalAlign="Left" Width="100px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoSafra" HeaderText="Safra" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
                    <HeaderTemplate>
                        Faixa de Comissão
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoFaixa" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAlterarFaixa" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirFaixa" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparFaixa" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorioFaixa" Text="Relatório" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjudaFaixa" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tabela:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblTabela" runat="server" Font-Bold="true" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Faixa Inicial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFaixaInicial" CssClass="txtDecimal" runat="server" data-ToolTip="default" ToolTip="" />
                            </div>
                            <div class="collbl">
                                Faixa Final:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFaixaFinal" CssClass="txtDecimal" runat="server" data-ToolTip="default" ToolTip="" />
                            </div>
                            <div class="collbl">
                                Indice:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtIndice" CssClass="txtDecimal" runat="server" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacao" runat="server" TextMode="MultiLine" Width="651px" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 250px;">
                            <asp:GridView ID="gridFaixaDeComissao" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="98%" OnSelectedIndexChanged="gridFaixaDeComissao_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True">
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:CommandField>
                                    <asp:BoundField HeaderText="Faixa Inicial" DataField="FaixaInicial" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Faixa Final" DataField="FaixaFinal" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Indice" DataField="Indice" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Observacoes" HeaderText="Observa&#231;&#227;o" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
