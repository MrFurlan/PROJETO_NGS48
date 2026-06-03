<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaNotaTroca.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaNotaTroca" %>
<script type="text/javascript">
    function pageLoadConsultaNotaTroca() {
        $("#btnFiltrar", "#divConsultaNotaTroca").button();
        $("#btnLimpar", "#divConsultaNotaTroca").button();
        $("#btnFechar", "#divConsultaNotaTroca").button();
    }

    $(document).ready(function () {
        pageLoadConsultaNotaTroca();
    });

    var prmConsultaNotaTroca = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaNotaTroca.add_endRequest(pageLoadConsultaNotaTroca);
</script>
<div id="divConsultaNotaTroca" class="uc" title="Consulta de Nota" style="display: none;" >
    <asp:UpdatePanel ID="updConsultaNotaTroca" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkCompra" Text="Consultar NF Dev. p/ Compra" runat="server" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkTransferencia" Text="Consultar NF Dev. p/ Transf. Tit." runat="server" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkTroca" Text="Consultar Notas De Troca" runat="server" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkCessaoDeCredito" Text="Consultar Notas De Cessao de Credito" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconSair" runat="server">
                            <asp:LinkButton ID="LnkSair" Text="Sair" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumNota" runat="server" ClientIDMode="Static" Width="100px" />
                </div>
                <div class="collbluc">
                    Ano:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumAno" runat="server" ClientIDMode="Static" Width="100px" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridNotas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridNotas_SelectedIndexChanged"
                    Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Movimento" HeaderText="Movimento" />
                        <asp:BoundField DataField="Empresa_Id" HeaderText="Empresa" />
                        <asp:BoundField DataField="EndEmpresa_Id" HeaderText="End" />
                        <asp:BoundField DataField="EmpresaNome" HeaderText="Nome" />
                        <asp:BoundField DataField="EmpresaCidade" HeaderText="Cidade" />
                        <asp:BoundField DataField="EmpresaEstado" HeaderText="Estado" />
                        <asp:BoundField DataField="Cliente_ID" HeaderText="Cliente" />
                        <asp:BoundField DataField="EndCliente_Id" HeaderText="End" />
                        <asp:BoundField DataField="Nome" HeaderText="Nome" />
                        <asp:BoundField DataField="Cidade" HeaderText="Cidade" />
                        <asp:BoundField DataField="Estado" HeaderText="Estado" />
                        <asp:BoundField DataField="EntradaSaida_Id" HeaderText="E/S" />
                        <asp:BoundField DataField="Nota_Id" HeaderText="Nota" />
                        <asp:BoundField DataField="Serie_Id" HeaderText="Serie" />
                        <asp:BoundField DataField="Produto_Id" HeaderText="Produto" />
                        <asp:BoundField DataField="NomeProduto" HeaderText="Nome" />
                        <asp:BoundField DataField="QtdeNota" HeaderText="Qtde" DataFormatString="{0:N4}"
                            HtmlEncode="False" />
                        <asp:BoundField DataField="SaldoTroca" HeaderText="Saldo" DataFormatString="{0:N4}"
                            HtmlEncode="False" />
                        <asp:BoundField DataField="Unidade" HeaderText="UN." />
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
