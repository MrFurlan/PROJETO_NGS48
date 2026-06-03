<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucRegistrosIcmsAjustaResumo.ascx.vb"
    Inherits="NGS.Web.UI.ucRegistrosIcmsAjustaResumo" %>
<style type="text/css">
    .collbluc
    {
        width: 123px;
    }
</style>
<div id="divRegistrosIcmsAjustaResumo" class="uc" title="Registro de Apuração do ICMS"
    style="display: none;">
    <asp:UpdatePanel ID="updpnlRegistrosIcmsAjustaResumo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconSair" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Código:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCodigo" runat="server" Width="630px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlCodigo_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Cod.Ajust.Icms:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCodigoAjusteICMSSpedFiscal" runat="server" Width="630px"
                        AutoPostBack="True" Font-Overline="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Descrição Ajust.Icms.:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescAjusteICMS" runat="server" TextMode="MultiLine" Width="620px"
                        ReadOnly="true" Height="48px" BackColor="#F7F6F3" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" TextMode="MultiLine" Width="620px"
                        MaxLength="300" Height="48px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Valor:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValor" CssClass="txtDecimal" runat="server" Style="text-align: right;" />
                </div>
            </div>
            <div class="bordagrid" style="height: 300px;">
                <asp:GridView ID="GridResumo" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridResumo_SelectedIndexChanged">
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <EditRowStyle BackColor="#999999" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descricao">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Valor" HeaderText="Valor">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="AjustesApuracaoIcms_Id" HeaderText="CodAjusteIcms">
                            <ItemStyle Width="50px" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
