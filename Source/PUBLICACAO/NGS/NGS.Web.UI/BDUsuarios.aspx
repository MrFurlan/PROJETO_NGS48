<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="BDUsuarios.aspx.vb" Inherits="NGS.Web.UI.BDUsuarios" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmBDUsuarios" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlBDUsuarios" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Banco de Dados
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
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Banco:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtBanco" runat="server" CssClass="texto" MaxLength="50" Width="540px"
                        data-ToolTip="default" ToolTip="Nome do banco." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Host do Servidor:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtHost" runat="server" CssClass="texto" MaxLength="50" Width="540px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Usuário do Banco:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtUsuarioBD" runat="server" CssClass="texto" MaxLength="50" Width="540px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Senha do Banco:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSenhaBD" runat="server" MaxLength="50" Width="540px" TextMode="Password" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grd" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grd_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True"
                            ItemStyle-HorizontalAlign="Center" ItemStyle-VerticalAlign="Middle" />
                        <asp:BoundField DataField="Banco_Id" HeaderText="Banco" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-VerticalAlign="Middle" />
                        <asp:BoundField DataField="HostServidor" HeaderText="Host do Servidor" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-VerticalAlign="Middle" />
                        <asp:BoundField DataField="UsuarioBanco" HeaderText="Usuário do Banco" HeaderStyle-HorizontalAlign="Left"
                            ItemStyle-VerticalAlign="Middle" />
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
