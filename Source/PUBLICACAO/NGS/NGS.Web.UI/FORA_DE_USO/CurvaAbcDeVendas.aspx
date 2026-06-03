<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CurvaAbcDeVendas.aspx.vb" Inherits="NGS.Web.UI.CurvaAbcDeVendas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCurvaAbcDeVendas" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCurvaAbcDeVendas" runat="server">
        <ContentTemplate>
            <table style="width: 100%; border: 0px none;">
                <tr>
                    <td class="titulotabela" colspan="2">
                        <label>
                            Curva ABC de Vendas
                        </label>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <table class="actions" style="width: 100%;">
                            <tr>
                                <td class="iconConsultar" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkConsultar" runat="server" CssClass="lnkMenu">
                                        <span>Consultar</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconLimpar" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="LinkLimpar" runat="server" CssClass="lnkMenu">
                                        <span>Limpar</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconAjuda" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkAjuda" runat="server" CssClass="lnkMenu">
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
                    <td style="width: 100px">
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Unidade:</span>
                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:DropDownList ID="DdlUnidade" runat="server" Width="600px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                            AutoPostBack="True" Font-Names="Courier New" Font-Size="8pt">
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
                        <asp:DropDownList ID="DdlEmpresa" runat="server" Width="600px" Font-Names="Courier New"
                            Font-Size="8pt">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Data:</span>
                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlMes" runat="server" Width="120px">
                        </asp:DropDownList>
                        &nbsp;De&nbsp;&nbsp;<asp:DropDownList ID="ddlAno" runat="server" Width="72px">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <div id="painel" class="bordasimples" style="min-height: 415px; max-height: 415px;
                            max-width: 100%;">
                            <asp:GridView ID="GridAbc" runat="server" AutoGenerateColumns="False" CssClass="gridview"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="QuantidadeNF" HeaderText="Qt.Nfs."></asp:BoundField>
                                    <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Valor"></asp:BoundField>
                                    <asp:BoundField DataField="ParcelaIndividual" HeaderText="%Faturamento">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorTotal" HeaderText="%Faturamento Total">
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
    </asp:UpdatePanel>
</asp:Content>
