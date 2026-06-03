<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaPedidosXNotas.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaPedidosXNotas" %>
<script type="text/javascript">
    function pageLoadConsultaPedidosXNotas() {
        $("#MainContent_ucConsultaPedidosXNotas_btnFechar", "#divConsultaPedidosXNotas").button();
    }

    $(document).ready(function () {
        pageLoadConsultaPedidosXNotas();
    });

    var prmConsultaPedidosXNotas = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaPedidosXNotas.add_endRequest(pageLoadConsultaPedidosXNotas);
</script>
<div id="divConsultaPedidosXNotas" class="uc" title="Consulta de Pedidos X Notas"
    style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaPedidosXNotas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="bordagrid">
                <asp:GridView ID="GridNotas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridNotas_SelectedIndexChanged"
                    OnRowDataBound="GridNotas_RowDataBound">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Movimento" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Empresa_id" HeaderText="Empresa">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="EndEmpresa_id" HeaderText="End">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="EndCliente" HeaderText="End">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="NomeCliente" HeaderText="Nome">
                            <HeaderStyle HorizontalAlign="Left" Width="350px" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ES" HeaderText="E/S">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Serie" HeaderText="S&#233;rie">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Nota" HeaderText="Nota">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Romaneio" HeaderText="Romaneio">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Operacao" HeaderText="OP">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SubOperacao" HeaderText="SO">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" DataFormatString="{0:n0}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Unitario" HeaderText="Unit&#225;rio" DataFormatString="{0:n6}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Valor" HeaderText="Valor" DataFormatString="{0:n2}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="DescSituacao" HeaderText="Situação">
                            <HeaderStyle />
                            <ItemStyle Font-Bold="True" />
                        </asp:BoundField>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:HiddenField ID="hdfNossaEmissao" runat="server" Value='<%# Eval("NossaEmissao") %>' />
                                <asp:LinkButton ID="lnkConsultar" runat="server" OnClick="lnkConsultar_Click">
                                                <img src='<%# ResolveUrl("~/Images/find.png") %>' alt="Consultar NF-e" data-ToolTip="default" title="Consultar NF-e" /> 
                                </asp:LinkButton>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="row">
                <div class="coltxt" style="float: right;">
                    <asp:Button ID="btnFechar" runat="server" Text="Fechar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
