<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucCancelamentoDeCheque.ascx.vb"
    Inherits="NGS.Web.UI.ucCancelamentoDeCheque" %>
<div id="divCancelamentoDeCheque" class="uc" title="Inutilização de Cheques" style="display: none;">
    <style type="text/css">
        .collbluc
        {
            width: 87px;
        }
    </style>
    <asp:UpdatePanel ID="updpnlCancelamentoDeCheques" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconSair" runat="server">
                            <asp:LinkButton ID="lnkSair" Text="Sair" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEmpresa" runat="server" Width="540px" Enabled="false" MaxLength="2" />
                    <asp:HiddenField ID="txtCondigoEmp" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Banco:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtBanco" runat="server" Enabled="false" Width="66px" />
                </div>
                <div class="collbluc">
                    Agencia:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAgencia" Enabled="false" runat="server" Width="66px" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDigitoAgencia" Enabled="false" runat="server" Width="30px" />
                </div>
                <div class="collbluc">
                    Conta:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtConta" runat="server" Enabled="false" Width="66px" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDigitoConta" runat="server" Enabled="false" Width="30px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Cheque Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtChequeInicial" runat="server" Width="66px" />
                </div>
                <div class="collbluc">
                    Cheque Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtChequeFinal" runat="server" Width="66px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Observação:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtObservacao" Class="TextLine" runat="server" Width="540px" MaxLength="200"
                        TextMode="MultiLine" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridCancelados" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="DataCancelamento" HeaderText="Data" ReadOnly="True" DataFormatString="{0:dd/MM/yyyy}">
                            <HeaderStyle HorizontalAlign="center" Width="66px" />
                            <ItemStyle HorizontalAlign="center" Width="66px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Pedido" HeaderText="Pedido" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="center" Width="50px" />
                            <ItemStyle HorizontalAlign="center" Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Titulo" HeaderText="Título" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="center" Width="50px" />
                            <ItemStyle HorizontalAlign="center" Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="UsuarioCancelamento" HeaderText="Usuário Cancelamento"
                            ReadOnly="True">
                            <HeaderStyle HorizontalAlign="center" Width="100px" />
                            <ItemStyle HorizontalAlign="center" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="TipoCancelamento" HeaderText="Tipo" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="center" Width="50px" />
                            <ItemStyle HorizontalAlign="center" Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Observacao" HeaderText="Observação" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
