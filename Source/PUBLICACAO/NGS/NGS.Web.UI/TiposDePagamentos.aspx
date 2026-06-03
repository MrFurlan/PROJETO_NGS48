<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="TiposDePagamentos.aspx.vb" Inherits="NGS.Web.UI.TiposDePagamentos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngTiposDePagamentos" runat="server">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlTiposDePagamentos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Tipos de Pagamentos
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
                    <asp:TextBox ID="txtCodigo" runat="server" Width="100px" data-ToolTip="default"
                        ToolTip="Código de cadastro do tipo de pagamento." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" Width="500px" data-ToolTip="default"
                        ToolTip="Descrição do tipo de pagamento." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Envia ao Banco:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbSim" runat="server" AutoPostBack="True" Text="Sim" GroupName="Banco" data-ToolTip="default" ToolTip="Envia para Banco." />
                    <asp:RadioButton ID="rbNao" runat="server" AutoPostBack="True" Text="Não" Checked="True" GroupName="Banco" data-ToolTip="default" ToolTip="Não envia para Banco." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Código Sefaz:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigoSefaz" runat="server" Width="100px" Enabled="false" data-ToolTip="default"
                        ToolTip="Código de cadastro do tipo de pagamento da Sefaz." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridTiposDePagamentos" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridTiposDePagamentos_SelectedIndexChanged">
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
                        <asp:BoundField DataField="EnviaAoBanco" HeaderText="Banco">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="TPagSefaz" HeaderText="C&#243;digo Sefaz">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
