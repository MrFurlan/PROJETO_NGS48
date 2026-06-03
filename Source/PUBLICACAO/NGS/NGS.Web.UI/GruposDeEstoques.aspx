<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="GruposDeEstoques.aspx.vb" Inherits="NGS.Web.UI.GruposDeEstoques" %>

<asp:Content ID="Scriptcontent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngGruposDeEstoques" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlGruposDeEstoques" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Grupos de Estoques
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
                    <asp:TextBox ID="txtCodigo" runat="server" TabIndex="1" Width="100px" data-ToolTip="default"
                        ToolTip="Código de cadastro do grupo." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" TabIndex="2" Width="500px" data-ToolTip="default"
                        ToolTip="Descrição do grupo." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Paramêtros:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkCusto" runat="server" Text="Custo" data-ToolTip="default" ToolTip="Marcar as opções de acordo com o grupo do estoque." />
                    <asp:CheckBox ID="chkAgrupaFinanceiro" runat="server" Text="Agrupa Financeiro" data-ToolTip="default"
                        ToolTip="Marcar as opções de acordo com o grupo do estoque." />
                    <asp:CheckBox ID="chkRelatorioEstoque" runat="server" Text="Relatório De Estoque"
                        data-ToolTip="default" ToolTip="Marcar as opções de acordo com o grupo do estoque." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridGruposDeEstoques" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridGruposDeEstoques_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" Width="25px" />
                        </asp:CommandField>
                        <asp:BoundField DataField="Codigo" HeaderText="Código">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" Width="60px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:CheckBoxField DataField="Custo" HeaderText="Custo">
                            <HeaderStyle HorizontalAlign="Right" Width="50px" />
                            <ItemStyle HorizontalAlign="Right" Width="50px" />
                        </asp:CheckBoxField>
                        <asp:CheckBoxField DataField="AgrupaFinanceiro" HeaderText="Agrupa Fin.">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                        </asp:CheckBoxField>
                        <asp:CheckBoxField DataField="RelatorioEstoque" HeaderText="Relatório De Estoque">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="70px" />
                        </asp:CheckBoxField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
