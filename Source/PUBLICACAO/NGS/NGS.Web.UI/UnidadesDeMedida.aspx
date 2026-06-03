<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="UnidadesDeMedida.aspx.vb" Inherits="NGS.Web.UI.UnidadesDeMedida" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngUnidadeDeMedida" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlUnidadeDeMedida" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Unidades de Medidas
            </div>
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
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
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
                    <asp:TextBox ID="txtCodigo" runat="server" Width="100px" Enabled="False" data-ToolTip="default"
                        ToolTip="Código de cadastro da unidade de medida." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" Width="500px" Enabled="False" data-ToolTip="default"
                        ToolTip="Descrição da unidade de medida." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Código Indea:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigoIndea" runat="server" Width="100px" Enabled="False" data-ToolTip="default"
                        ToolTip="Código de registro no Instituto de Defesa Agropecuária do Estado de Mato Grosso." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridUnidadesDeMedidas" runat="server" AutoGenerateColumns="False"
                    CellPadding="3" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridUnidadesDeMedidas_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="25px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Unidade" HeaderText="C&#243;digo" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="UnidadeIndea" HeaderText="Indea" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
