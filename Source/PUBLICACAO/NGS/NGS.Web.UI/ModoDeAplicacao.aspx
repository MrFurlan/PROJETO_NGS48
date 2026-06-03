<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ModoDeAplicacao.aspx.vb" Inherits="NGS.Web.UI.ModoDeAplicacao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngModeDeAplicacao" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlModoDeAplicacao" runat="server">
        <ContentTemplate>
            <table style="width: 100%; border: 0px none;">
                <tr>
                    <td class="titulotabela" colspan="2">
                        <label>
                            Modo de Aplicação
                        </label>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <table class="actions" style="width: 100%;">
                            <tr>
                                <td class="iconNovo" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkNovo" runat="server">
                                        <span>Gravar</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconAtualizar" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkAtualizar" runat="server">
                                        <span>Atualizar</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconExcluir" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkExcluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;">
                                        <span>Excluir</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconLimpar" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkLimpar" runat="server">
                                        <span>Limpar</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconAjuda" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkAjuda" runat="server">
                                        <span>Ajuda</span>
                                    </asp:LinkButton>
                                </td>
                                <td style="display: block;">
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="width: 130px">
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">CulturaXPragaXFito:</span>
                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlCulturaPragaFito" runat="server" Enabled="False" 
                             Width="505px">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Forma de Aplicacao:</span>
                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlFormaDeAplicacao" runat="server" Enabled="False" 
                             Width="505px">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td style="vertical-align: top">
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Descrição:</span>
                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:TextBox ID="txtDescricao" runat="server" Enabled="False" Height="60px" TextMode="MultiLine"
                            Width="500px"/>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:Panel ID="pnlModoDeAplicacao" runat="server" Height="200px" Width="100%">
                            <asp:GridView ID="gridModoDeAplicacao" runat="server" CellPadding="4" ForeColor="#333333"
                                GridLines="None" Width="98%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            </asp:GridView>
                        </asp:Panel>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
