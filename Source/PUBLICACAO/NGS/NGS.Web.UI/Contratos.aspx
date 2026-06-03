<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Contratos.aspx.vb" Inherits="NGS.Web.UI.Contratos"
    MasterPageFile="~/Principal.Master" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function Arquivo() {
            $("#btnAdicionar").click();
        }
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngContratos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlContrato" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Contratos
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
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" runat="server" Text="Relatório" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Código:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" runat="server" Enabled="false" TabIndex="2" Width="500px" data-ToolTip="default"
                        ToolTip="Código do Contrato." />
                </div>
            </div>

            <div class="row" runat="server" visible="true">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" TabIndex="2" Width="500px" data-ToolTip="default"
                        ToolTip="Descrição do Contrato." />
                </div>
            </div>
            <div class="row">

                <div class="collbl">
                    Arquivo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeDoArquivo" runat="server" Enabled="false" Style="width: 200px;" />
                    <asp:FileUpload ID="fupArquivo" OnChange="Arquivo()" runat="server" Width="120px" Font-Size="11px" ClientIDMode="Static" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridContratos" runat="server" AutoGenerateColumns="False" ForeColor="#333333"
                    GridLines="None" Width="100%">
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
                            <HeaderStyle HorizontalAlign="Left" Width="210px" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                            <HeaderStyle HorizontalAlign="Left" Width="600px" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>

                        <asp:TemplateField HeaderText="Arquivo">
                            <ItemTemplate>
                                <asp:ImageButton ID="imgDownload" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/download32x32.png" Style="margin-top: 0;"
                                    Height="22px" Width="22px" OnClick="imgDownload_Click" ToolTip="Baixar Arquivo" />
                            </ItemTemplate>
                        </asp:TemplateField>

                        <asp:TemplateField HeaderText="Excluir">
                            <ItemTemplate>
                                <asp:ImageButton ID="imgExcluir" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                    OnClick="imgExcluir_Click" data-ToolTip="default" ToolTip="Excluir"
                                    OnClientClick="return confirm('Deseja realmente excluir o produto?');" />
                            </ItemTemplate>
                            <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                        </asp:TemplateField>

                    </Columns>
                </asp:GridView>
            </div>
            <asp:Button ID="btnAdicionar" runat="server" Text="Adicionar" ClientIDMode="Static"
                OnClick="btnAdicionar_Click" CssClass="none" />
            <asp:Label ID="lblMsg" runat="server" />
        </ContentTemplate>
        <Triggers>
            <asp:PostBackTrigger ControlID="btnAdicionar" />
        </Triggers>
    </asp:UpdatePanel>
</asp:Content>
