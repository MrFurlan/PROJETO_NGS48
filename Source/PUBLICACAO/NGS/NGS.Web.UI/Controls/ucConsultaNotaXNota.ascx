<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaNotaXNota.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaNotaXNota" %>
<script type="text/javascript">
    function pageLoadConsultaNotaXNota() {
        $("#btnFechar", "#divConsultaNotaXNota").button();
    }

    $(document).ready(function () {
        pageLoadConsultaNotaXNota();
    });

    var prmConsultaNotaXNota = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaNotaXNota.add_endRequest(pageLoadConsultaNotaXNota);
</script>
<div id="divConsultaNotaXNota" class="uc" title="Consulta de Nota X Nota" style="display: none;">
    <center>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr>
                <td>
                    <asp:UpdatePanel ID="updpnlNotaXNota" runat="server">
                        <ContentTemplate>
                            <table class="borda" width="100%">
                                <tr>
                                    <td colspan="3" class="titulotabela">
                                        Notas X Notas
                                    </td>
                                </tr>
                                <tr>
                                    <td class="rotulo" align="right">
                                        Nota Fiscal:
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtNotaFiscal" runat="server"/>
                                    </td>
                                    <td align="right">
                                        <asp:Button ID="btnConsultar" runat="server" Text="Consultar" CssClass="botao" UseSubmitBehavior="False"
                                            
                                            
                                            OnClick="btnConsultar_Click" />&nbsp;&nbsp;&nbsp;
                                        <asp:Button ID="btnLimpar" runat="server" Text="Limpar" CssClass="botao" UseSubmitBehavior="False"
                                            
                                            
                                            OnClick="btnLimpar_Click" />&nbsp;&nbsp;&nbsp;
                                        <input id="btnSair" 
                                            class="botao" 
                                            onclick="parent.location.href='Expedicao.aspx';" type="button" value="Sair" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3">
                                        <asp:GridView ID="gridNxN" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                            ForeColor="#333333" GridLines="None" Width="100%">
                                            <EditRowStyle BackColor="#999999" />
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <Columns>
                                                <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Nota" HeaderText="Nota">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Movimento" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Descricao" HeaderText="Tipo">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Size="Small" Text="X"/>
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="OrigemEmpresa" HeaderText="Empresa">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="OrigemCliente" HeaderText="Cliente">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="OrigemNota" HeaderText="Nota">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="OrigemMovimento" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="OrigemDescricao" HeaderText="Tipo">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                            </Columns>
                                        </asp:GridView>
                                    </td>
                                </tr>
                            </table>
                        </ContentTemplate>
                    </asp:UpdatePanel>
                </td>
            </tr>
        </table>
    </center>
</div>
