<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="AlteraEmpresaPedido.aspx.vb" Inherits="NGS.Web.UI.AlteraEmpresaPedido" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function selectAll(chkAll) {
            var chk = $('#' + chkAll.id);
            var checked = chk.attr('checked') == "checked";

            $(".chk input[type='checkbox']", "#<%=gridPedidos.ClientID%>").not(".chkAll").each(function () {
                $(this).attr("checked", checked);
            });
        };

        function downloadArquivo() {
            alert("Processo Concluido");
            __doPostBack("", "DownloadArquivo");
        };
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngAletarEmpresaPedido" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAletarEmpresaPedido" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Alterar Empresa do(s) Pedido(s).
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server" visible="false">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" runat="server" Text="Gravar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade Origem:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="650px" 
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa Origem:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" Enabled="false" runat="server" Width="611px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultaCliente" runat="server" Text="&gt;" CssClass="btn" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Width="150px" data-tooltip="default" ToolTip="Para informar mais de um pedido, separe-os com vírgula(,)." />
                </div>
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData1" CssClass="calendario" runat="server" Width="80px" />
                </div>
                <div class="coltxt">
                    á
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData2" runat="server" CssClass="calendario" Width="80px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlOperacao" runat="server" Width="650px" AutoPostBack="true"/>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Sub-Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSubOperacao" runat="server" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="650px" />
                </div>
            </div>
            <div class="row">
                <asp:Panel ID="Panel1" runat="server" GroupingText="Empresa Destino" 
                    Font-Size="Medium">
                    <div class="row">
                        <div class="collbl">
                            Unidade:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlUnidadeDestino" runat="server" Width="650px" 
                                Enabled="false" AutoPostBack="True" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Empresa:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlEmpresaDestino" runat="server" Width="650px" Enabled="false" />
                        </div>
                    </div>
                </asp:Panel>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridPedidos" runat="server" CellPadding="4" AutoGenerateColumns="False"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                <input id="chkAll" onclick="selectAll(this);" type="checkbox" />
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkSelecionado" CssClass="chk" runat="server" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Pedido" HeaderText="Pedido" />
                        <asp:BoundField DataField="Nome" HeaderText="Cliente" />
                        <asp:BoundField DataField="Observacoes" HeaderText="Observações" />
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:CheckBox ID="chktemAut" runat="server" Checked='<%# Eval("temAut") %>' Enabled="False" Text="Autorização" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:CheckBox ID="chktemCom" runat="server" Checked='<%# Eval("temCom") %>' Enabled="False" Text="Comissão" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:CheckBox ID="chktemTro" runat="server" Checked='<%# Eval("temTro") %>' Enabled="False" Text="Troca" />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
