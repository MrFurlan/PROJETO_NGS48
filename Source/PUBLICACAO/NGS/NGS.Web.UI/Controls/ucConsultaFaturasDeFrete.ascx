<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaFaturasDeFrete.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaFaturasDeFrete" %>
<script type="text/javascript">
    function pageLoadConsultaFaturasDeFrete() {
        $("#btnFechar", "#divConsultaFaturasDeFrete").button();
    }

    $(document).ready(function () {
        pageLoadConsultaFaturasDeFrete();
    });

    var prmConsultaFaturasDeFrete = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaFaturasDeFrete.add_endRequest(pageLoadConsultaFaturasDeFrete);
</script>
<div id="divConsultaFaturasDeFrete" class="uc" title="Consulta de Faturas de Frete"
    style="display: none;">
    <asp:UpdatePanel ID="updConsultaFaturasDeFrete" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="bordagrid">
                <asp:GridView ID="DgFaturas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="DgFaturas_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ShowSelectButton="True" ButtonType="Button" SelectText="&gt;">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                        </asp:CommandField>
                        <asp:BoundField DataField="CodigoNota" HeaderText="CT-e">
                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoFatura" HeaderText="Fatura">
                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoConveniado" HeaderText="Conveniado">
                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ConveniadoNome" HeaderText="Nome">
                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento">
                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorDaFatura" DataFormatString="{0:n2}" HeaderText="Valor Fatura">
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorLancadoFatura" DataFormatString="{0:n2}" HeaderText="Valor Lan&#231;ado">
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorSaldoFatura" DataFormatString="{0:n2}" HeaderText="Saldo">
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoTitulo" HeaderText="Titulo" Visible="False">
                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoEmpresa" HeaderText="Empresa" Visible="False">
                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="EndEmpresa" HeaderText="EndEmpresa" Visible="False">
                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ConveniadoEnd" HeaderText="EndConveniado" Visible="False">
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Encargo" HeaderText="Encargo">
                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="row">
                <div class="painelright">
                    <asp:Button ID="btnFechar" runat="server" ClientIDMode="Static" Text="Fechar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
