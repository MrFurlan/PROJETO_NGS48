<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="IngredienteAtivo.aspx.vb" Inherits="NGS.Web.UI.IngredienteAtivo" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function confirmarAlteracao() {
            if (confirm("Confirma alteração da Nota Fiscal?")) {
                __doPostBack("", "Alterar");
            }
        }

        function confirmarExclusao() {
            if (confirm("Confirma exclusão da Nota Fiscal?")) {
                __doPostBack("", "Excluir");
            }
        }
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngIngredienteAtivo" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlIngredienteAtivo" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Ingrediente Ativo
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
                    <asp:TextBox ID="txtCodigo" runat="server" Enabled="False" data-ToolTip="default" ToolTip="Número de cadastro do ingrediente ativo." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Estado Físico:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEstadoFisico" runat="server" Width="511px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo Químico:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupoQuimico" runat="server" Width="511px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nome Comum:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeComum" runat="server" Width="500px" MaxLength="50" data-ToolTip="default" ToolTip="Nome simplificado do produto." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nome Químico:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeQuimico" runat="server" Width="500px" MaxLength="100" data-ToolTip="default" ToolTip="Nome técnico do produto." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Solubilidade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSolubilidade" runat="server" Width="500px" TextMode="MultiLine" data-ToolTip="default" ToolTip="Informações para dissolver o produto." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Peso Molecular:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPesoMolecular" runat="server" Width="500px" MaxLength="50" data-ToolTip="default" ToolTip="" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Ponto Fusão:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPontoFusao" runat="server" Width="500px" MaxLength="100" data-ToolTip="default" ToolTip="" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pressão Vapor:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPressaoVapor" runat="server" Width="500px" MaxLength="100" data-ToolTip="default" ToolTip="" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridIngredienteAtivo" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridIngredienteAtivo_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                            <HeaderStyle Width="50px" />
                            <ItemStyle Width="50px" />
                        </asp:CommandField>
                        <asp:BoundField DataField="CodigoIA" HeaderText="C&#243;digo" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="110px" />
                            <ItemStyle HorizontalAlign="Left" Width="110px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoEstadoFisico" HeaderText="Estado Fisico" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoGrupoQuimico" HeaderText="Grupo Qu&#237;mico">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="NomeComum" HeaderText="Nome Comum" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="NomeQuimico" HeaderText="Nome Qu&#237;mico" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
