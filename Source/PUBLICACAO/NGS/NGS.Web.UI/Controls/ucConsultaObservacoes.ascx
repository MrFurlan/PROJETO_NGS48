<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaObservacoes.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaObservacoes" %>
<script type="text/javascript">
    function pageLoadConsultaObservacoes() {
        $("#btnFechar", "#divConsultaObservacoes").button();
    }

    $(document).ready(function () {
        pageLoadConsultaObservacoes();
    });

    var prmConsultaObservacoes = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaObservacoes.add_endRequest(pageLoadConsultaObservacoes);
</script>
<div id="divConsultaObservacoes" class="uc" title="Consulta de Observações" style="display: none;">
    <asp:UpdatePanel ID="updConsultaObservacoes" runat="server" UpdateMode="Conditional">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="bordagrid">
                <asp:GridView ID="GridObservacoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                            <HeaderStyle Width="30px" />
                            <ItemStyle Width="30px" />
                        </asp:CommandField>
                        <asp:BoundField DataField="Codigo_Id" HeaderText="C&#243;d.">
                            <HeaderStyle Width="50px" HorizontalAlign="Left" />
                            <ItemStyle Width="50px" HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Estado" HeaderText="UF">
                            <HeaderStyle Width="30px" HorizontalAlign="Left" />
                            <ItemStyle Width="30px" HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
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
