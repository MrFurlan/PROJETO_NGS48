<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ClassesDeOperacoes.aspx.vb" Inherits="NGS.Web.UI.ClassesDeOperacoes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl
        {
            width: 140px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmClassesDeOperacoes" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlClassesDeOperacoes" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Classes de Operações
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
                    <asp:TextBox ID="txtCodigo" TabIndex="1" runat="server" Width="150px" data-ToolTip="default"
                        ToolTip="Número de  cadastro  das classes de operação." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" TabIndex="2" runat="server" Width="500px" data-ToolTip="default"
                        ToolTip="Descrição de  cadastro  das classes de operação." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classificação de Frete:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlClassificacaoFrete" runat="server" Width="502px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtro:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkOperacao" runat="server" Text="Carregar no Cadastro de Operações"
                        data-ToolTip="default" ToolTip="Marcar os filtros desejados." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkSubOperacao" runat="server" Text="Carregar no Cadastro de SubOperação"
                        data-ToolTip="default" ToolTip="Marcar os filtros desejados." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridClassesDeOperacoes" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridClassesDeOperacoes_SelectedIndexChanged">
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
                        <asp:TemplateField HeaderText="Operação" ItemStyle-HorizontalAlign="Center" ItemStyle-VerticalAlign="Middle">
                            <EditItemTemplate>
                                <asp:CheckBox ID="CheckBox1" runat="server" Checked='<%# Bind("Operacao") %>' />
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="CheckBox1" runat="server" Checked='<%# Bind("Operacao") %>' Enabled="false" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="SubOperação" ItemStyle-HorizontalAlign="Center" ItemStyle-VerticalAlign="Middle">
                            <EditItemTemplate>
                                <asp:CheckBox ID="CheckBox2" runat="server" Checked='<%# Bind("SubOperacao") %>' />
                            </EditItemTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="CheckBox2" runat="server" Checked='<%# Bind("SubOperacao") %>'
                                    Enabled="false" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="ClassificacaoDeFrete" HeaderText="Classificação De Frete">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
