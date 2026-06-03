<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PedidosEntreguesDeCompras.aspx.vb" Inherits="NGS.Web.UI.PedidosEntreguesDeCompras" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngPedidosEntreguesDeCompras" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPedidosEntreguesDeCompras" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Compras Faturadas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" runat="server">
                                <span>Consultar</span>
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
                                    <asp:LinkButton ID="lnkExcelCC" runat="server" Text="Excel CC" />
                                </li>
                            </ul>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server">
                                <span>Limpar</span>
                            </asp:LinkButton>
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="600px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged" data-ToolTip="default"
                        ToolTip="Unidade de Negocio Empresarial." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="600px" data-ToolTip="default"
                        ToolTip="Empresa Para negociação." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="112px"
                        data-ToolTip="default" ToolTip="Data final a inicial do pedido." />
                </div>
                <div class="collbl">
                    à
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="112px"
                        data-ToolTip="default" ToolTip="Data final a inicial do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="568px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="ddlCliente" OnClick="ddlCliente_Click" runat="server" Text="&gt;"
                        CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
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
                <div class="coltxt lg">
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
            <div class="row" id="divCentroDeCusto" runat="server" Visible="false">
                <div class="collbl">
                    Centro de Custo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCentroDeCusto" runat="server" Width="600px" data-ToolTip="default"
                        ToolTip="Selecionar o Centro de Custo." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Visualizar:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadCliente" runat="server" AutoPostBack="True" Text="Por Cliente" GroupName="Visualizar"
                        Checked="True" data-ToolTip="default" ToolTip="Selecionar o modo de visualização." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadProduto" runat="server" AutoPostBack="True" Text="Por Produto" GroupName="Visualizar"
                        data-ToolTip="default" ToolTip="Selecionar o modo de visualização." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadCentroDeCusto" runat="server" AutoPostBack="True" Text="Por Centro de Custo" GroupName="Visualizar"
                        data-ToolTip="default" ToolTip="Selecionar o modo de visualização - Apenas para EXCEL CC." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ckApenasFinanceiro" runat="server" Text="Apenas Financeiro" data-ToolTip="default"
                        ToolTip="Selecionar o modo de visualização." />
                </div>
            </div>
            <table style="width: 100%; border: 0px none;">
                <tr>
                    <td colspan="2">
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
                                            Height="15px" ImageUrl="~/images/ico-mais.gif" data-ToolTip="default" ToolTip="Adicionar CFOP para Lista">
                                        </asp:ImageButton><br />
                                        <br />
                                        <asp:ImageButton ID="imgRemover" OnClick="imgRemover_Click" runat="server" Width="15px"
                                            Height="15px" ImageUrl="~/images/ico-menos.gif" data-ToolTip="default" ToolTip="Remover CFOP da Lista">
                                        </asp:ImageButton>
                                    </td>
                                    <td>
                                        <asp:Panel ID="pnlSelecionados" runat="server" Width="100%" Height="130px">
                                            <asp:ListBox ID="lstCfopSelecionados" runat="server" Height="100%" Width="100%">
                                            </asp:ListBox>
                                        </asp:Panel>
                                    </td>
                                </tr>
                            </table>
                        </asp:Panel>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" runat="server">
                        <div class="bordasimples" runat="server" style="min-height: 415px; max-height: 415px;
                            max-width: 100%; overflow: auto;">
                            <asp:GridView ID="GridProduto" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <ItemStyle Width="25px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Data" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data">
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nota" HeaderText="Nota"></asp:BoundField>
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
                                    <asp:BoundField DataField="Quantidade" DataFormatString="{0:N0}" HeaderText="Quantidade">
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Unitario" DataFormatString="{0:N}" HeaderText="Unitario">
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Valor"></asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" runat="server">
                        <div class="bordasimples" runat="server" style="min-height: 415px; max-height: 415px;
                            max-width: 100%; overflow: auto;">
                            <asp:GridView ID="GridCliente" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <ItemStyle Width="25px"></ItemStyle>
                                    </asp:CommandField>
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
                                    <asp:BoundField DataField="Data" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Emiss&#227;o">
                                    </asp:BoundField>
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
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
