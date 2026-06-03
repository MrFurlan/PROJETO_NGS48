<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Cfop.aspx.vb" Inherits="NGS.Web.UI.Cfop" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmCfop" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCfop" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Códigos Fiscais
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel1">
                    <HeaderTemplate>
                        Título
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
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodTitulo" runat="server" MaxLength="4" OnTextChanged="txtCodTitulo_TextChanged"
                                    Style="width: 150px" data-ToolTip="default" ToolTip="Código de cadastro do CFOP." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDesTitulo" runat="server" MaxLength="300" Style="width: 579px;"
                                    data-ToolTip="default" ToolTip="Descrição do CFOP." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="grdAbaTitulo" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grdAbaTitulo_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Código" />
                                    <asp:BoundField DataField="Descricao" HeaderText="Descrição" />
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel2">
                    <HeaderTemplate>
                        Código Fiscal
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoCfop" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizarCfop" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirCfop" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparCfop" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorioCfop" Text="Relatório" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="LinkButton6" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtReduzidoCfop" runat="server" MaxLength="4" Enabled="false" Width="150px"
                                    data-ToolTip="default" ToolTip="Código de cadastro da situação tributária." />
                                <asp:Label ID="lblTitulo" runat="server" Font-Bold="True" ForeColor="Red" Style="display: inline;
                                    width: 400px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricaoCfop" runat="server" MaxLength="300" Style="width: 563px;"
                                    data-ToolTip="default" ToolTip="Descrição da situação tributária." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="grdAbaFiscal" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="grdAbaFiscal_SelectedIndexChanged"
                                Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Código" />
                                    <asp:BoundField DataField="Descricao" HeaderText="Descrição" />
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
</asp:Content>
