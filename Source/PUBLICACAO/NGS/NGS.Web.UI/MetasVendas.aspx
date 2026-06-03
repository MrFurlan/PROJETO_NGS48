<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="MetasVendas.aspx.vb" Inherits="NGS.Web.UI.MetasVendas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmMetasVendas" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlMetasVendas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <table style="width: 100%; border: 0px none;">
                <tr>
                    <td class="titulotabela">
                        <label>
                            Metas de Vendas
                        </label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                            <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                                <HeaderTemplate>
                                    Metas
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <table style="width: 100%; border: 0px none;">
                                        <tr>
                                            <td colspan="3">
                                                <div class="menu_acoes">
                                                    <div class="acoes">
                                                        <ul>
                                                            <li runat="server">
                                                                <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                                                            </li>
                                                        </ul>
                                                    </div>
                                                </div>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="width: 100px;">
                                                <div class="headerGray">
                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Empresa:</span>
                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                </div>
                                            </td>
                                            <td style="width: 480px">
                                                <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                                                <asp:TextBox ID="txtEmpresa" runat="server" Width="500px" />
                                            </td>
                                            <td>
                                                <asp:Button ID="btnConsultaEmpresa" runat="server" Text=">" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div class="headerGray">
                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Ano:</span>
                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                </div>
                                            </td>
                                            <td style="width: 401px">
                                                <asp:DropDownList ID="ddlAno" Width="85px" runat="server" />
                                            </td>
                                            <td style="text-align: right;">
                                                <asp:Button CssClass="botao" ID="btnNovo" Text="Nova Meta" runat="server" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="3">
                                                <div class="bordasimples" style="min-height: 415px; max-height: 415px; max-width: 100%;">
                                                    <asp:GridView ID="GridMetas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                        ForeColor="#333333" GridLines="None" Width="100%">
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
                                                            <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                                                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Left" Width="50px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="Safra_Id" HeaderText="Safra">
                                                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Left" Width="70px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="Cultura_Id" HeaderText="Cultura">
                                                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                            </asp:BoundField>
                                                        </Columns>
                                                    </asp:GridView>
                                                </div>
                                            </td>
                                        </tr>
                                    </table>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                                <HeaderTemplate>
                                    Metas 2
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <table style="width: 100%; border: 0px none;">
                                        <tr>
                                            <td colspan="2">
                                                <div class="menu_acoes">
                                                    <div class="acoes">
                                                        <ul>
                                                            <li runat="server">
                                                                <asp:LinkButton class="iconAjuda" ID="AjudaMeta2" runat="server" Text="Ajuda" />
                                                            </li>
                                                        </ul>
                                                    </div>
                                                </div>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="width: 100px;">
                                                <div class="headerGray">
                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Ano:</span>
                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                </div>
                                            </td>
                                            <td>
                                                <asp:DropDownList ID="ddlAno2" runat="server" Width="85px">
                                                </asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div class="headerGray">
                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Empresa:</span>
                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                </div>
                                            </td>
                                            <td>
                                                <asp:TextBox ID="txtEmpresa2" runat="server" Width="500px" />
                                                <asp:Button ID="btnConsultaEmpresa2" runat="server" Text=">" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div class="headerGray">
                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Cotação:</span>
                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                </div>
                                            </td>
                                            <td>
                                                <asp:TextBox ID="txtCotacao" runat="server" Width="500px" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div class="headerGray">
                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Safra:</span>
                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                </div>
                                            </td>
                                            <td>
                                                <table style="width: 100%;">
                                                    <tr>
                                                        <td style="width: 200px;">
                                                            <asp:DropDownList ID="ddlSafra" Width="150px" runat="server">
                                                            </asp:DropDownList>
                                                        </td>
                                                        <td style="width: 100px;">
                                                            <div class="headerGray">
                                                                <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Cultura:</span>
                                                                <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                            </div>
                                                        </td>
                                                        <td style="width: 155px;">
                                                            <asp:DropDownList ID="ddlCultura" Width="150px" runat="server">
                                                            </asp:DropDownList>
                                                        </td>
                                                        <td>
                                                            <asp:Button ID="btnAdicionar" CssClass="botao" runat="server" Text="+" Width="40px" />
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="2">
                                                <div class="bordasimples" style="min-height: 415px; max-height: 415px; max-width: 100%;">
                                                    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                        ForeColor="#333333" GridLines="None" Width="100%">
                                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <EditRowStyle BackColor="#999999" />
                                                        <Columns>
                                                            <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                                                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Left" Width="50px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="Safra" HeaderText="Safra">
                                                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Left" Width="70px"></ItemStyle>
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="Cultura" HeaderText="Cultura">
                                                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                            </asp:BoundField>
                                                        </Columns>
                                                    </asp:GridView>
                                                </div>
                                            </td>
                                        </tr>
                                    </table>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                        </ajaxToolkit:TabContainer>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="Server" />
</asp:Content>
