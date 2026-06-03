<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="NaturezaDeRendimentos.aspx.vb" Inherits="NGS.Web.UI.NaturezaDeRendimentos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngNaturezaDeRendimentos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlNaturezaDeRendimentos" runat="server">
        <ContentTemplate>
            <div id="dialog-confirm" title="ATENÇÃO?" style="display: none;">
                <p><span class="ui-icon ui-icon-alert" style="float: left; margin: 0 7px 20px 0;"></span>Tem certeza que deseja excluir este registro?</p>
            </div>
            <div class="titulodiv">
                Natureza De Rendimentos
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
                    <label id="resultado">Código:</label>
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" CssClass="txtNumerico5" MaxLength="5" runat="server" Width="50px"
                        data-tooltip="default" ToolTip="Código de registro." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" data-tooltip="default" TextMode="MultiLine" ToolTip="Descrição da Natureza de Rendimento."
                        runat="server" Width="800px" MaxLength="500" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Típo da Pessoa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlPessoa" TabIndex="3" runat="server" Width="100px">
                        <asp:ListItem Selected="True" Value="0">Nenhuma</asp:ListItem>
                        <asp:ListItem Value="1">Fisica</asp:ListItem>
                        <asp:ListItem Value="2">Juridica</asp:ListItem>
                        <asp:ListItem Value="3">Ambas</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridNaturezaDeRendimentos" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" radius-border="5px" ForeColor="#333333" GridLines="None"
                    Width="100%" OnSelectedIndexChanged="GridNaturezaDeRendimentos_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle CssClass="hStyleBordagrid" />
                    <SelectedRowStyle BackColor="#e1e7ef" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                            <HeaderStyle HorizontalAlign="Left" Width="25px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="25px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Codigo" HeaderText="Código">
                            <HeaderStyle HorizontalAlign="Left" Width="60px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="TipoPessoa" HeaderText="Tipo da Pessoa" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="60px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
