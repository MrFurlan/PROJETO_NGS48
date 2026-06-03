<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ClasseToxicologica.aspx.vb" Inherits="NGS.Web.UI.ClasseToxicologica" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngClasseToxicologica" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlClasseToxicologica" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Classe Toxicológica
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
                    Código:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" runat="server" data-ToolTip="default" ToolTip="Número de cadastro da classe." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" Width="500px" data-ToolTip="default" ToolTip="Descrição do grau de intoxicação." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cor:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeCor" runat="server" data-ToolTip="default" ToolTip="Cor que identifica o grau toxicológico." />
                </div>
                <div class="collbl">
                    Cor Hex:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCor" runat="server" data-ToolTip="default" ToolTip="Código da cor que identifica o grau toxicológico." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridClasseToxicologica" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnPreRender="gridClasseToxicologica_PreRender"
                    OnSelectedIndexChanged="gridClasseToxicologica_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                            <HeaderStyle Width="50px" />
                            <ItemStyle Width="50px" />
                        </asp:CommandField>
                        <asp:BoundField DataField="codigo" HeaderText="C&#243;digo" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="100px" />
                            <ItemStyle HorizontalAlign="Left" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="descricao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Cor" HeaderText="Cor" />
                        <asp:BoundField DataField="CodigoCor" HeaderText="Cor Hex." />
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
