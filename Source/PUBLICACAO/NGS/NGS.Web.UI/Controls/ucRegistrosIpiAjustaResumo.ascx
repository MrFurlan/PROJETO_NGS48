<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucRegistrosIpiAjustaResumo.ascx.vb"
    Inherits="NGS.Web.UI.ucRegistrosIpiAjustaResumo" %>
<div id="divRegistrosIpiAjustaResumo" class="uc" title="Consulta de Vencimentos"
    style="display: none;">
    <style type="text/css">
        .collbl
        {
            width: 135px;
        }
    </style>
    <asp:UpdatePanel ID="updpnlRegistrosIpiAjustaResumo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="subtitulodiv">
                Registro de Apuração do IPI
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" runat="server" Text="Gravar" />
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" runat="server" Text="Excluir" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li class="iconSair" runat="server">
                            <asp:LinkButton ID="lnkFechar" runat="server" Text="Fechar" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Código:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlDescricaoRAIPI" runat="server" Width="496px" Font-Size="9pt"
                        OnSelectedIndexChanged="DdlDescricaoRAIPI_SelectedIndexChanged" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cod.Ajust.IPI:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlDescricaoAjusteIPISpedFiscal" runat="server" AutoPostBack="True"
                        Font-Size="9pt" Width="496px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Origem Doc. Vinculado:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlOrgDocVinc" runat="server" AutoPostBack="True" Font-Size="9pt"
                        Width="496px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" TextMode="MultiLine" Width="488px"
                        Height="23px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Valor:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValor" runat="server" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridResumo" runat="server" AutoGenerateColumns="False" BackColor="White"
                    BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="2" OnSelectedIndexChanged="GridResumo_SelectedIndexChanged"
                    Width="100%">
                    <FooterStyle BackColor="White" ForeColor="#000066" />
                    <RowStyle ForeColor="#000066" />
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
                        <asp:BoundField DataField="AjustesApuracaoIPI_Id" HeaderText="Cod.Ajuste IPI" />
                        <asp:BoundField DataField="OrigemDocumentoAjusteIpi_Id" HeaderText="Cod. Origem Docto" />
                    </Columns>
                    <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                    <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
