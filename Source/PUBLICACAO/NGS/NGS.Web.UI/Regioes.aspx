<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Regioes.aspx.vb" Inherits="NGS.Web.UI.Regioes" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="ScriptContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRegioes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Regiões
            </div>
            <ajaxToolkit:TabContainer Style="margin-top: 4px;" ID="tcRegioes" runat="server"
                ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabRegiao" runat="server">
                    <HeaderTemplate>
                        Região
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
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server"
                                            OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
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
                                <asp:TextBox ID="txtCodigo" TabIndex="1" runat="server" Width="100px" data-ToolTip="default"
                                    ToolTip="Código de cadastro da região." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricao" TabIndex="2" runat="server" Width="500px" data-ToolTip="default"
                                    ToolTip="Descrição do local." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="GridRegioes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridRegioes_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <ItemStyle Width="25px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabMicroRegiao" runat="server">
                    <HeaderTemplate>
                        Micro Região
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoMicroRegiao" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizarMicroRegiao" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirMicroRegiao" runat="server" Text="Excluir" 
                                            OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparMicroRegiao" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorioMicroRegiao" Text="Relatório" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjudaMicroRegiao" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Região:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblRegiao" runat="server" Width="400px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigoMicroRegiao" TabIndex="1" runat="server" Width="100px"
                                    Enabled="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricaoMicroRegiao" TabIndex="2" runat="server" Width="500px"
                                    data-ToolTip="default" ToolTip="Descrição do local." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="grdMicroRegiao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <ItemStyle Width="25px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
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
