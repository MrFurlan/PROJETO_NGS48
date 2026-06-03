<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PedidosEntreguesDeVendas.aspx.vb" Inherits="NGS.Web.UI.PedidosEntreguesDeVendas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadPedidosEntreguesDeVenda() {
            $("#MainContent_lstTipoDoItem").multiselect();
        }

        $(document).ready(function () {
            pageLoadPedidosEntreguesDeVenda();
            var prmPedidosEntreguesDeVenda = Sys.WebForms.PageRequestManager.getInstance();
            prmPedidosEntreguesDeVenda.add_endRequest(pageLoadPedidosEntreguesDeVenda);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPedidosEntreguesDeVendas" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPedidosEntreguesDeVendas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Vendas Faturadas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" runat="server">
                                <span>Consultar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server">
                                <span>Limpar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcelD" runat="server" Text="Excel Dados" />
                                </li>
                            </ul>
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server">
                                <span>Ajuda</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="600px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" data-ToolTip="default" ToolTip="Unidade de Negocio Empresarial." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkEmpresa" ToolTip="Unificar o cnpj das empresas." CssClass="multiselect" Text="Empresa:" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="600px" data-ToolTip="default" ToolTip="Empresa Para negociação/consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="112px" data-ToolTip="default" ToolTip="Data inicial a final do pedido." />
                </div>
                <div class="collbl">
                    à:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="112px" data-ToolTip="default" ToolTip="Data inicial a final do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="562px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdCliente" OnClick="cmdCliente_Click" runat="server" Text="&gt;"
                        CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo do Item:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstTipoDoItem" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="602px" data-ToolTip="default" ToolTip="Marcar o item para pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadCliente" runat="server" Text=" Por Cliente " Checked="True"
                        GroupName="Visualizar" data-ToolTip="default" ToolTip="Filtrar por cliente ou produto." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadProduto" runat="server" Text=" Por Produto " GroupName="Visualizar" data-ToolTip="default" ToolTip="Filtrar por cliente ou produto." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ckApenasFinanceiro" runat="server" Checked="True" Text="Apenas Financeiro" data-ToolTip="default" ToolTip="Marcar para selecionar somente o financeiro." />
                </div>
                <div class="collbl">
                    Estado Físico:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEstadoFisico" runat="server" Width="185px" ToolTip="Estado Físico." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbOperacao" runat="server" Width="294px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbOperacao_SelectedIndexChanged" />
                    <asp:DropDownList ID="cmbSubOperacao" runat="server" Width="295px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 100%; margin-bottom: 4px;">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo CFOP:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGruposCFOP" runat="server" Width="600px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlGruposCFOP_SelectedIndexChanged" data-ToolTip="default"
                        ToolTip="Selecionar o tipo do CFOP." />
                </div>
            </div>
            <div class="row">
                <asp:Panel ID="Panel3" runat="server" Width="700px" Height="132px">
                    <table style="width: 100%; border: 0px none;">
                        <tr>
                            <td>
                                <asp:Panel ID="pnlCfop" runat="server" Width="100%" Height="130px">
                                    <asp:ListBox ID="lstCfop" runat="server" Height="100%" Width="100%"></asp:ListBox>
                                </asp:Panel>
                            </td>
                            <td style="width: 30px; text-align: center;">
                                <asp:ImageButton ID="imgAdicionar" OnClick="imgAdicionar_Click" runat="server" Width="15px"
                                    Height="15px" ImageUrl="~/images/ico-mais.gif" data-ToolTip="default" ToolTip="Adicionar CFOP para Lista"></asp:ImageButton><br />
                                <br />
                                <asp:ImageButton ID="imgRemover" OnClick="imgRemover_Click" runat="server" Width="15px"
                                    Height="15px" ImageUrl="~/images/ico-menos.gif" data-ToolTip="default" ToolTip="Remover CFOP da Lista"></asp:ImageButton>
                            </td>
                            <td>
                                <asp:Panel ID="pnlSelecionados" runat="server" Width="100%" Height="130px">
                                    <asp:ListBox ID="lstCfopSelecionados" runat="server" Height="100%" Width="100%"></asp:ListBox>
                                </asp:Panel>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
            </div>
            <div class="bordagrid" runat="server">
                <asp:GridView ID="GridProduto" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Data" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data"></asp:BoundField>
                        <asp:BoundField DataField="Nota" HeaderText="Nota"></asp:BoundField>
                        <asp:BoundField DataField="NCM" HeaderText="NCM" />
                        <asp:BoundField DataField="Produto" HeaderText="Produto">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Unidade" HeaderText="Unidade">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="NomeDoProduto" HeaderText="NomeDoProduto">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Quantidade" DataFormatString="{0:N0}" HeaderText="Quantidade"></asp:BoundField>
                        <asp:BoundField DataField="Unitario" DataFormatString="{0:N}" HeaderText="Unitario"></asp:BoundField>
                        <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Valor"></asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="bordagrid" runat="server">
                <asp:GridView ID="GridCliente" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Cidade" HeaderText="Cidade">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Bairro" HeaderText="Bairro">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Estado" HeaderText="UF">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Nota" HeaderText="Nota"></asp:BoundField>
                        <asp:BoundField DataField="Data" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Emiss&#227;o"></asp:BoundField>
                        <asp:BoundField DataField="Quantidade" DataFormatString="{0:N0}" HeaderText="Quantidade">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Valor">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
