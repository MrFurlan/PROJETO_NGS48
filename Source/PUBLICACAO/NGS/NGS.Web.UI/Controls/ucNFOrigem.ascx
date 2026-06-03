<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucNFOrigem.ascx.vb"
    Inherits="NGS.Web.UI.ucNFOrigem" %>
<div id="divNFOrigem" class="uc" title="Informe a nota fiscal de origem" style="display: none;">
    <asp:UpdatePanel ID="updpnlNFOrigem" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkGravar" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconFechar" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoClienteFrete" runat="server" />
                    <asp:TextBox ID="txtClienteFrete" runat="server" Enabled="False" Font-Names="monospace"
                        Width="592px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdClienteFrete" runat="server" OnClick="cmdConsultaClienteFrete_Click"
                        Text="&gt;" UseSubmitBehavior="False" CssClass="btn" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataIniFrete" runat="server" CssClass="calendario"
                        Width="88px" />
                </div>
                <div class="coltxt">
                    à
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFimFrete" runat="server" CssClass="calendario"
                        Width="88px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Número:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtNotaXContabilizacao" runat="server" />
                    <asp:TextBox ID="txtNumNFFrete" runat="server" CssClass="texto" TabIndex="7"
                        Width="88px" />
                </div>
                <div class="collbluc">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerieNFFrete" runat="server" CssClass="txtUpper" MaxLength="3"
                        TabIndex="8" Width="36px" />
                </div>
                <div class="collbluc">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdEntradas" runat="server" Text="Entradas" GroupName="Notas" data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                    <asp:RadioButton ID="rdSaidas" runat="server" Text="Saidas" Checked="True" GroupName="Notas" data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                </div>
            </div>
            <div class="bordagrid" style="height: 195px;">
                <asp:GridView ID="grdNotaFreteConsulta" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField ShowHeader="False">
                            <ItemTemplate>
                                <asp:CheckBox ID="chkNotaXNota" runat="server" />
                            </ItemTemplate>
                            <ItemStyle Width="30px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="Cliente_Id" HeaderText="Cliente">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Deposito" HeaderText="Dep&#243;sito" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Destino" HeaderText="Destino" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Operacao" HeaderText="Operação">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="NumNota" HeaderText="Nota" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Tipo" HeaderText="Tipo" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Quantidade" HeaderText="Quantidade">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Valor" HeaderText="Valor">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
