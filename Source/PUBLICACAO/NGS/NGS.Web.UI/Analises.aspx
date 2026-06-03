<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Analises.aspx.vb" Inherits="NGS.Web.UI.Analises" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngAnalises" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAnalises" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Tabela de Análises
            </div>
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
                    Código Analise:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtAnaliseId" MaxLength="10" Width="200px" CssClass="txtNumerico" data-ToolTip="default"
                        ToolTip="Código de cadastro das análises." />
                </div>
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtDescricao" MaxLength="50" Width="200px" data-ToolTip="default"
                        ToolTip="Nome das análises." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Índice Mínimo:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtIndiceMinimo" MaxLength="30" Width="200px" data-ToolTip="default"
                        ToolTip="Informe o índice mínimo para esta análise." />
                </div>
                <div class="collbl">
                    Índice Máximo:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtIndiceMaximo" MaxLength="30" Width="200px" data-ToolTip="default"
                        ToolTip="Informe o índice máximo para esta análise." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Opções:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtOpcao" MaxLength="100" Width="400px" data-ToolTip="default"
                        ToolTip="Informe as opcoes separanco com '-' para valor e descricao e ';' para multiplas opcoes. ex: 0-NAO;1-SIM;2-Positivo ou NAO;SIM;POSITIVO no primeiro caso grava 0 e no segundo NAO." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GrdAnalises" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GrdAnalises_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                            <HeaderStyle Width="25px" />
                        </asp:CommandField>
                        <asp:BoundField DataField="Codigo" HeaderText="Código" HtmlEncode="False" FooterText="Código">
                            <FooterStyle HorizontalAlign="Left"></FooterStyle>
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="200px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Análise" FooterText="Análise">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="IndiceMinimo" HeaderText="Índice Mínimo" HtmlEncode="False"
                            FooterText="Índice Mínimo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="IndiceMaximo" HeaderText="Índice Máximo" HtmlEncode="False"
                            FooterText="Índice Máximo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Opcao" HeaderText="Opções" HtmlEncode="False"
                            FooterText="Opções">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
