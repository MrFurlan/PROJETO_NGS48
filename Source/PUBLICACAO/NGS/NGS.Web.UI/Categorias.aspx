<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Categorias.aspx.vb" Inherits="NGS.Web.UI.Categorias" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrm" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCategorias" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Categorias
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
                                OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" runat="server" Text="Relatório" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Código:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" runat="server" Width="100px" data-ToolTip="default" ToolTip="Código de cadastro da categoria do cliente." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" Width="500px" data-ToolTip="default"
                        ToolTip="Descrição da categoria do cliente." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Funrural:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFunrural" runat="server" Width="100px" />
                </div>
                <div class="collbl">
                    Base Pis Cofins:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtBasePisCofins" runat="server" Width="100px" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridCategorias" runat="server" AutoGenerateColumns="False" OnSelectedIndexChanged="GridCategorias_SelectedIndexChanged"
                    OnPreRender="GridCategorias_PreRender" CellPadding="4" ForeColor="#333333" GridLines="None"
                    Width="100%">
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
                        <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Funrural" DataFormatString="{0:N}" HeaderText="Funrural">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right" Width="60px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="BasePisCofins" DataFormatString="{0:N}" HeaderText="Cofins">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right" Width="60px"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
