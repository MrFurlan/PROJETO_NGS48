<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Grupos.aspx.vb" Inherits="NGS.Web.UI.Grupos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngGrupos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlGrupos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Grupos de Usuários
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
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtGrupo" TabIndex="1" runat="server" Width="250px" data-ToolTip="default"
                        ToolTip="Identificação do grupo de usuário." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" TabIndex="2" runat="server" Width="500px" data-ToolTip="default"
                        ToolTip="Nome do grupo de usuário." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridGruposDeUsuarios" runat="server" Width="100%" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="GridUsuarios_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="25px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Grupo" HeaderText="C&#243;digo" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="150px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
