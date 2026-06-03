<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PrecoDePauta.aspx.vb" Inherits="NGS.Web.UI.PrecoDePauta" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPrecoDePauta" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPrecoDePauta" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Preço de Pauta
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
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Estados:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEstado" runat="server" Width="445px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox CssClass="calendario" ID="txtData" runat="server" ReadOnly="true" data-ToolTip="default"
                        ToolTip="Data de atualização do preço de pauta." />
                </div>
                <div class="collbl">
                    Preço:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPreco" CssClass="txtDecimal10" runat="server" ReadOnly="True"
                        data-ToolTip="default" ToolTip="Valor." />
                </div>
            </div>
            <div class="painelleft" style="width: 59.5%;">
                <div class="bordagrid">
                    <asp:GridView ID="gridProduto" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridProduto_SelectedIndexChanged">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <EditRowStyle BackColor="#999999" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                <HeaderStyle Width="20px" />
                                <ItemStyle Width="20px" />
                            </asp:CommandField>
                            <asp:BoundField DataField="Codigo" HeaderText="Produto" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" Width="50px" />
                                <ItemStyle HorizontalAlign="Left" Width="50px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Nome" HeaderText="Nome" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <div class="painelright" style="width: 39.5%;">
                <div class="bordagrid">
                    <asp:GridView ID="gridPrecoDePauta" runat="server" CellPadding="4" ForeColor="#333333"
                        GridLines="None" AutoGenerateColumns="False" Width="100%" OnSelectedIndexChanged="gridPrecoDePauta_SelectedIndexChanged">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <EditRowStyle BackColor="#999999" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                <HeaderStyle Width="20px" />
                                <ItemStyle Width="20px" />
                            </asp:CommandField>
                            <asp:BoundField DataField="CodigoEstado" HeaderText="Estado" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" Width="50px" />
                                <ItemStyle HorizontalAlign="Left" Width="50px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Data" HeaderText="Data" HtmlEncode="False" DataFormatString="{0:dd/MM/yyyy}">
                                <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                <ItemStyle HorizontalAlign="Left" Width="100px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Preco" HeaderText="Preço" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" Width="100px" />
                                <ItemStyle HorizontalAlign="Right" Width="100px" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
