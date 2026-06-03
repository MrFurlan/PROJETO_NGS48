<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ContasSaldosDisponiveis.aspx.vb" Inherits="NGS.Web.UI.ContasSaldosDisponiveis" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmContasSaldosDisponiveis" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlContasSaldosDisponiveis" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Contas de Saldos Disponíveis
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Incluir" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                    <asp:TextBox ID="txtEmpresa" runat="server" Enabled="False" Width="450px" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgEmpresa" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                        OnClick="imgEmpresa_Click" data-ToolTip="default" ToolTip="Empresa Para negociação/consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaInicial" runat="server" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgContaInicial" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                        OnClick="imgContaInicial_Click" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeContaInicial" runat="server" Enabled="False" Width="315px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaFinal" runat="server" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgContaFinal" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                        OnClick="imgContaFinal_Click" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeContaFinal" runat="server" Enabled="False" Width="315px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" Width="450px" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridContasSaldosDisponiveis" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridContasSaldosDisponiveis_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Empresa_Id" HeaderText="Empresa">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="EndEmpresa_Id" HeaderText="End">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ContaInicial_Id" HeaderText="Conta Inicial">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ContaFinal_Id" HeaderText="Conta Final">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
</asp:Content>
