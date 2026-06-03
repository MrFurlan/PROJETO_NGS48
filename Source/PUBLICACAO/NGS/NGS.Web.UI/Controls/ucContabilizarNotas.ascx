<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucContabilizarNotas.ascx.vb" Inherits="NGS.Web.UI.ucContabilizarNotas" %>
<script type="text/javascript">

</script>

<div id="divContabilizarNotas" class="uc" title="Contabilizar Notas" style="display: none;">
    <asp:UpdatePanel ID="updpnlContabilizarNotas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConfirmar" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkConfirmar" runat="server">
                                <span>Confirmar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconLimpar" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server">
                                <span>Limpar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconSair" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkFechar" runat="server">
                                <span>Fechar</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresas:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEmpresas" Width="600px" runat="server" data-tooltip="default" ToolTip="CNPJs das Empresas." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Clientes:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClientes" Width="600px" runat="server" data-tooltip="default" ToolTip="CNPJs dos clientes." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Incial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server"
                        Width="126px" data-ToolTip="default" ToolTip="Data inicial do intervalo." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server"
                        Width="126px" data-ToolTip="default" ToolTip="Data final do intervalo." />
                </div>
                <div class="collbl">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEntradaSaida" Width="27px" runat="server" MaxLength="1" data-ToolTip="default" ToolTip="Entrada ou Saida." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    OperaçãoXEstado:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtOperacaoXEstado" runat="server" Width="230px" data-ToolTip="default" ToolTip="Digite as IDs." />
                </div>
                <div class="collbluc">
                    Pedidos:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedidos" runat="server" Width="230px" data-tooltip="default" ToolTip="Número dos pedidos." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Notas:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNotas" width="600px" runat="server" data-tooltip="default" ToolTip="Digite o número das notas." />
                </div>
            </div>
            <asp:GridView ID="gridNotas" runat="server" Title="Notas Encontradas" AutoGenerateColumns="False" CellPadding="4"
                ForeColor="#333333" GridLines="None" Width="100%">
                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EditRowStyle BackColor="#999999" />
                <Columns>
                    <asp:BoundField DataField="Cliente_Id" HeaderText="Cliente">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundField>
                    <asp:BoundField DataField="Nota_Id" HeaderText="Nota">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundField>
                    <asp:BoundField DataField="Serie_Id" HeaderText="Serie">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundField>
                    <asp:BoundField DataField="EntradaSaida_Id" HeaderText="E/S">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundField>
                </Columns>
            </asp:GridView>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>