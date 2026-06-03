<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Autorizantes.aspx.vb" Inherits="NGS.Web.UI.Autorizantes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngAutorizantes" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="udppnlAutorizantes" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Autorizantes
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
                    Autorizante:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlAutorizantes" TabIndex="3" runat="server" Width="550px"
                        OnSelectedIndexChanged="DdlGrupos_SelectedIndexChanged" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Substituto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlSubstitutos" TabIndex="4" runat="server" Width="550px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nível:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlNivel" runat="server" Width="50px">
                        <asp:ListItem>1</asp:ListItem>
                        <asp:ListItem>2</asp:ListItem>
                        <asp:ListItem>3</asp:ListItem>
                        <asp:ListItem>9</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="collbl">
                    Cotação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlCotacao" runat="server" Width="50px">
                        <asp:ListItem Value="S">SIM</asp:ListItem>
                        <asp:ListItem Value="N">NÃO</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridAutorizantes" runat="server" AutoGenerateColumns="False" CellPadding="3"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridView1_SelectedIndexChanged">
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
                        <asp:BoundField DataField="Autorizante" HeaderText="Autorizante">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="200px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Substituto" HeaderText="Substituto">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="200px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Nivel" HeaderText="Nivel">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Cotacao" HeaderText="Cotac&#227;o">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
