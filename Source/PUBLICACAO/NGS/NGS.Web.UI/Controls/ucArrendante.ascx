<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucArrendante.ascx.vb"
    Inherits="NGS.Web.UI.ucArrendante" %>
<script type="text/javascript">
    function pageLoadConsultaArrendante() {
        $("#btnFechar", "#divConsultaArrendantes").button();
    }

    $(document).ready(function () {
        pageLoadConsultaArrendante();
    });

    var prmConsultaArrendante = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaArrendante.add_endRequest(pageLoadConsultaArrendante);
</script>
<div id="divConsultaArrendantes" class="uc" title="Consulta de Arrendantes" style="display: none;">
    <asp:UpdatePanel ID="updpnlArrendante" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" type="submit" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconConfirmar" runat="server">
                            <asp:LinkButton ID="lnkConfirmar" Text="Confirmar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
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
                    <asp:HiddenField ID="txtCodigoArrendante" runat="server" />
                    <asp:TextBox ID="txtArrendante" runat="server" Enabled="False" Width="480px" required />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnArrendante" runat="server" OnClick="btnArrendante_Click" Text=">"
                        CssClass="btn" UseSubmitBehavior="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Data do Contrato:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataContrato" CssClass="calendario" runat="server" AutoPostBack="True"
                        OnTextChanged="txtDataContrato_TextChanged" />
                </div>
                <div class="collbluc">
                    Vencimento:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataVencimento" CssClass="calendario" runat="server" AutoPostBack="True"
                        OnTextChanged="txtDataVencimento_TextChanged" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridMatricula" runat="server" CellPadding="4" ForeColor="#333333"
                    GridLines="None" Width="100%" AutoGenerateColumns="False">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="Matricula_Id" HeaderText="Matricula">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Area" HeaderText="Área">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="AreaArrendada" HeaderText="Arrendada">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Saldo" HeaderText="Saldo">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Arrendamento">
                            <ItemTemplate>
                                <asp:TextBox ID="txtArrendamento" runat="server" CssClass="txtDecimal" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Observacao">
                            <ItemTemplate>
                                <asp:TextBox ID="txtObservacao" runat="server" Width="557px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
