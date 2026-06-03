<%@ Page Title="Conversão de Unidades" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="UnidadeConversao.aspx.vb" Inherits="NGS.Web.UI.UnidadeConversao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngUnidadeConversao" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlUnidadeConversao" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Conversão de Unidades
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
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" />
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
                    Unidade de Origem:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeOrigem" runat="server" AutoPostBack="true" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade de Destino:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDestino" runat="server" AutoPostBack="true" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Fator:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFator" class="txtDecimal6" runat="server" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridUnidadeConversao" runat="server" AutoGenerateColumns="False"
                    CellPadding="3" ForeColor="#333333" GridLines="None" Width="100%">
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
                        <asp:BoundField DataField="CodigoUnidadeOrigem" HeaderText="Un. Origem" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoUnidadeDestino" HeaderText="Un. Destino" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Fator">
                            <ItemTemplate>
                                <asp:TextBox ID="txtFator" runat="server" Enabled="False" Text='<%# Bind("Fator", "{0:N6}")%>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" Width="120px" />
                            <ItemStyle HorizontalAlign="Right" Width="120px" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
