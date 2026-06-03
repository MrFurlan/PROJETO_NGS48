<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Reclassificacao.aspx.vb" Inherits="NGS.Web.UI.Reclassificacao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngReclassificacao" runat="server">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlReclassificacao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Reclassificação
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
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
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="590px" />
                </div>
            </div>

            <div class="painelright" style="width: 685px;" runat="server" visible="false">
                <div class="bordagrid" style="height: 100%;">
                    <asp:GridView ID="gridDescontos" runat="server" CellPadding="4" ForeColor="#333333"
                        GridLines="None" Width="98%" AutoGenerateColumns="False">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Font-Names="Tahoma" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Font-Names="Tahoma" />
                        <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" Font-Names="Tahoma" />
                        <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" Font-Names="Tahoma" />
                        <EditRowStyle BackColor="#999999" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:BoundField DataField="CodigoAnalise" />
                            <asp:BoundField DataField="Descricao" HeaderText="Descrição" />
                            <asp:TemplateField HeaderText="Percentual">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtPercentual" runat="server" Width="98%" CssClass="txtDecimal txt"
                                        Enabled="false" Text='<%# IIf(Convert.ToString(Eval("Percentual")) = "0", "",  Eval("Percentual", "{0:N2}"))%>' />
                                    <asp:DropDownList ID="ddlOpcao" runat="server" Visible="False" Width="98%" CssClass="txt">
                                    </asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Indice">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtIndice" runat="server" Width="98%" Enabled="false" Text='<%# Eval("Indice")%>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Desconto">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtDesconto" runat="server" Width="98%" Enabled="false" Text='<%# Eval("Desconto")%>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </div>
                <div class="row">
                    <div class="coltxt">
                        <asp:Button ID="cmdCalcular" runat="server" CssClass="botao" Enabled="False" Text="Calcular"
                            OnClick="cmdCalcular_Click" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="cmdAjustar" runat="server" CssClass="botao" Visible="false" Text="Ajustar"
                            OnClick="cmdAjustar_Click" />
                    </div>
                </div>
            </div>
            <div class="painelleft">
                <div class="row">
                    <div class="collbl">
                        Laudo:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtLaudo" runat="server" CssClass="texto" Enabled="False" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Romaneio:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtRomaneio" runat="server" CssClass="texto" Enabled="False" />
                    </div>
                </div>

                <div class="row">
                    <div class="collbl">
                        1a_Pesagem:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPrimeiraPesagem" runat="server" Enabled="False" CssClass="txtDecimal" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        2a_Pesagem:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtSegundaPesagem" runat="server" Enabled="False" CssClass="txtDecimal" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Peso Bruto:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPesoBruto" runat="server" Enabled="False" CssClass="txtDecimal" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Desconto:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDesconto" runat="server" Enabled="False" CssClass="txtDecimal" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Liquido:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtLiquido" runat="server" Enabled="False" CssClass="txtDecimal" />
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
