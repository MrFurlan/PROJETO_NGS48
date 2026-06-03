<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="TabelasDePrecos.aspx.vb" Inherits="NGS.Web.UI.TabelasDePrecos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmnTabelasDePrecos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlTabelasDePrecos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Tabelas de Preços
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" runat="server" Text="Gravar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server" Text="Atualizar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" runat="server" Text="Excluir"
                                OnClientClick="return msgconfirm(this);" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Código
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" runat="server" Width="50px" minlength="2" CssClass="txtNumerico"
                        data-tooltip="default" ToolTip="Codigo do preço" />
                </div>
                <div class="collbl">
                    Ativo:
                </div>
                <div class="coltxt" style="margin-left: 40px">
                    <asp:CheckBox ID="chkAtivo" runat="server" AutoPostBack="True" Text="Ativo"
                        data-ToolTip="default" ToolTip="Desativar ou ativar ID." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" data-tooltip="default" ToolTip="Nome de registro do preço. " runat="server" Width="445px" MaxLength="50" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridTabelasDePrecos" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" radius-border="5px" ForeColor="#333333" GridLines="None"
                    Width="100%" OnSelectedIndexChanged="GridTabelasDePrecos_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle CssClass="hStyleBordagrid" />
                    <SelectedRowStyle BackColor="#e1e7ef" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Codigo" HeaderText="Codigo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="100px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
