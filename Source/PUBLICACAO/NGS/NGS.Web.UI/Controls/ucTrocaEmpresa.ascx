<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucTrocaEmpresa.ascx.vb"
    Inherits="NGS.Web.UI.ucTrocaEmpresa" %>
<script type="text/javascript">
    function pageLoadConsultaEmpresa() {
        $("#btnFechar", "#divConsultaEmpresas").button();
    }

    $(document).ready(function () {
        pageLoadConsultaEmpresa();
    });

    var prmConsultaEmpresa = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaEmpresa.add_endRequest(pageLoadConsultaEmpresa);
</script>
<div id="divTrocaEmpresa" class="uc" title="Consulta de Empresas" style="display: none;">
    <asp:UpdatePanel ID="updConsultaEmpresa" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="painelleft" style="width: 20%;">
                <div class="bordagrid" style="height: 215px">
                    <asp:ListBox ID="LbxUnidades" runat="server" AutoPostBack="True" Style="width: 100%;
                        height: 100%; font-size: 12px; border: none; cursor: pointer;" />
                </div>
            </div>
            <div class="painelright" style="width: 79%;">
                <div class="bordagrid" style="height: 215px;">
                    <asp:GridView ID="GridClientes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%" OnRowCommand="GridClientes_RowCommand">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:TemplateField ShowHeader="False">
                                <ItemTemplate>
                                    <asp:Button ID="cmdSelecao" runat="server" CausesValidation="False" CommandArgument='<%# Eval("Codigo") & ";" & Eval("CodigoEndereco") %>'
                                        CommandName="Select" Text=">" UseSubmitBehavior="False" />
                                </ItemTemplate>
                                <ItemStyle Height="19px" Width="19px" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="Reduzido" HeaderText="C&#243;digo">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Nome" HeaderText="Nome " HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Cidade" HeaderText="Cidade" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="CodigoEstado" HeaderText="UF">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>
                            <asp:BoundField DataField="CodigoFormatado" HeaderText="CNPJ/CPF">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
                <div class="row">
                    <div class="coltxt" style="float: right;">
                        <asp:Button ID="btnFechar" runat="server" ClientIDMode="Static" UseSubmitBehavior="False"
                            CssClass="botao" Text="Fechar" />
                    </div>
                </div>
            </div>
        </ContentTemplate>
        <Triggers>
            <asp:AsyncPostBackTrigger ControlID="lnkTrocaEmpresa" />
        </Triggers>
    </asp:UpdatePanel>
</div>
