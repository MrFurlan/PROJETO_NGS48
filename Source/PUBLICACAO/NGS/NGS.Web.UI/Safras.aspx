<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Safras.aspx.vb" Inherits="NGS.Web.UI.Safras" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngSafras" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlSafras" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Safras
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" runat="server">
                                <span>Gravar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkAtualizar" runat="server">
                                <span>Atualizar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;">
                                <span>Excluir</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server">
                                <span>Limpar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" runat="server">
                                <span>Relatório</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server">
                                <span>Ajuda</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSafra" Style="text-transform: uppercase;" runat="server" Width="555px"
                        data-ToolTip="default" ToolTip="Preencher com o nome da nova safra." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlGrupo" runat="server" Width="560px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlGrupo_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlProduto" runat="server" Width="560px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período de:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtInicioSafra" CssClass="calendario" runat="server" Width="80px"
                        data-ToolTip="default" ToolTip="Data inicial ao final da safra." />
                </div>
                <div class="coltxt">
                    à
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtVencimento" CssClass="calendario" runat="server" Width="80px"
                        data-ToolTip="default" ToolTip="Data inicial ao final da safra." />
                </div>
                <div class="collbl">
                    Taxa:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTaxa" CssClass="txtDecimal" runat="server" Width="80px" data-ToolTip="default"
                        ToolTip="Taxa de juro." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Observação:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtObservacao" TextMode="MultiLine" runat="server" Width="555px"
                        data-ToolTip="default" ToolTip="Campo para inserir observação quando necessário." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridSafras" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridSafras_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                            <ItemStyle Width="25px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Safra" HeaderText="Safra">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Grupo" HeaderText="Grupo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Produto" HeaderText="Produto">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="InicioDeSafra" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Inicio Safra" />
                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Taxa" DataFormatString="{0:N}" HeaderText="Taxa">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Observacao" HeaderText="Observa&#231;&#227;o">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
