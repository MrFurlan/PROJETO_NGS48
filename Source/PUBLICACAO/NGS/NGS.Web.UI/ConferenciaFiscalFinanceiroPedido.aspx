<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="ConferenciaFiscalFinanceiroPedido.aspx.vb" Inherits="NGS.Web.UI.ConferenciaFiscalFinanceiroPedido" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadConferenciaFiscalFinanceiroPedido() {
            $("#MainContent_lstClasseOp").multiselect().multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadConferenciaFiscalFinanceiroPedido();
            var prmConferenciaFiscalFinanceiroPedido = Sys.WebForms.PageRequestManager.getInstance();
            prmConferenciaFiscalFinanceiroPedido.add_endRequest(pageLoadConferenciaFiscalFinanceiroPedido);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngCFFP" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCFFP" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Conferência de Fiscal/Financeiro do Pedido.
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbUnidadeNegocio" runat="server" AutoPostBack="True" Width="600px"
                        data-ToolTip="default" ToolTip="Unidade de Negocio Empresarial." />
                </div>
            </div>

            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." runat="server"
                        Text="Empresa:" data-ToolTip="default" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente."
                        runat="server" Text="Cliente:" data-ToolTip="default" />
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="570px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultaCliente" CssClass="btn" runat="server" Text="&gt;" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe Operacao:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstClasseOp" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="370px" data-ToolTip="default" ToolTip="Selecionar a classe de opeação desejada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="240px" />
                </div>
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData1" Width="80px" runat="server" CssClass="calendario" />
                </div>
                <div class="coltxt">
                    á
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData2" Width="80px" runat="server" CssClass="calendario" />
                </div>
            </div>
            <div class="row">
                <uc:SelecaoProduto id="ucSelecaoProduto" runat="server" />
            </div>
            <div class="row">
                <div class="collbl">
                    Filtro(s):
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkFiscalAberto" runat="server" Text="Fiscal Aberto" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkFinanceiroAberto" runat="server" Text="Financeiro Aberto" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridPedido" runat="server" CellPadding="4" AutoGenerateColumns="False"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="Pedido" HeaderText="Pedido" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="100px" />
                            <ItemStyle HorizontalAlign="Left" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Safra" HeaderText="Safra" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>


                        <asp:BoundField DataField="Antecipada" HeaderText="Antecipada" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>

                        <asp:BoundField DataField="Recompra" HeaderText="Recompra" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>

                        <asp:BoundField DataField="Troca" HeaderText="Troca" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>






                        <asp:BoundField DataField="FiscalAberto" HeaderText="Fiscal Aberto" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="FinanceiroAberto" HeaderText="Financeiro Aberto" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:LinkButton ID="lnkExtrato" runat="server" OnClick="lnkExtrato_Click" data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido">
                                    <img alt="Extrato" src="App_Themes/ngssolucoes/imagens/icon_consultar.png" />
                                </asp:LinkButton>

                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" Width="30px" />
                            <ItemStyle HorizontalAlign="Center" Width="30px" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
