<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="ExtratoDeBens.aspx.vb" Inherits="NGS.Web.UI.ExtratoDeBens" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngExtratoDeBens" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlExtratoDeBens" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Extrato de Bens
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server"
                                OnClick="lnkConsultar_Click" Text="Consultar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" runat="server"
                                OnClick="lnkRelatorio_Click" Text="Relatório" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server"
                                OnClick="lnkLimpar_Click" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server"
                                OnClick="lnkAjuda_Click" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="638px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo Ativo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupoDeAtivo" runat="server" Width="638px" />
                </div>
            </div>
            <%-- <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkPeriodo" runat="server" Text="Período de:" />
                </div>
                <div runat="server" visible="false">
                    <div class="coltxt">
                        <asp:TextBox ID="txtData1" CausesValidation="" CssClass="calendario" runat="server" Width="86px" />
                    </div>
                    <div class="coltxt">
                        á:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtData2" CssClass="calendario" runat="server" Width="86px" />
                    </div>
                </div>
            </div>--%>
            <div class="bordagrid">
                <asp:GridView ID="gridAtivos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:ImageButton ID="imgExtratoDeBens" runat="server" CssClass="imgconsultar" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                    OnClick="imgExtratoDeBens_Click" data-ToolTip="default" ToolTip="Visualizar o extrato do bem selecionado" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="DescricaoBem" HeaderText="Descricao do Bem">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DataAquisicao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Aquisi&#231;&#227;o">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorOriginal" DataFormatString="{0:N}" HeaderText="Valor">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Baixado" HeaderText="Baixado">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
