<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="PedidoOnsoft.aspx.vb" Inherits="NGS.Web.UI.PedidoOnsoft" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPedidoOnsoft" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPedidoOnsoft" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="1" Width="100%">
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabIntegrado" ID="TabIntegrado">
                    <HeaderTemplate>
                        Pedidos Integrados OnMobile
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="titulodiv">
                            Pedidos OnMobile
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
                                <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeNegocio_SelectedIndexChanged"
                                    Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                                <asp:TextBox ID="txtClientes" runat="server" Width="557px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdBuscaCliente" OnClick="cmdBuscaCliente_Click" runat="server" Text=">"
                                    CssClass="btn" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default"
                                    ToolTip="Selecionar o cliente desejado." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Período:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data inicial do lançamento." />
                                &nbsp;à
                                <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data final do lançamento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedido" runat="server" data-ToolTip="default" ToolTip="Número do pedido no NGS." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Situação:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdBloqueado" runat="server" Text="Bloqueado" Checked="True" GroupName="Situacao" data-ToolTip="default" ToolTip="Pedidos que não foram Liberados." />
                                <asp:RadioButton ID="rdLiberado" runat="server" Text="Liberado" GroupName="Situacao" data-ToolTip="default" ToolTip="Pedidos que já foram Liberados." />
                                <asp:RadioButton ID="rdTodos" runat="server" Text="Todos" GroupName="Situacao" data-ToolTip="default" ToolTip="Todos os Pedidos." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridPedidoOnsoft" CssClass="gridSort" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" Wrap="True" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="PedidoNumPedCli" HeaderText="Número NGS">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="PedidoNum" HeaderText="Número OnMobile">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="PedidoData" HeaderText="Data">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="PedidoVrpagar" DataFormatString="{0:n2}" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ClienteCod" HeaderText="Cliente Cod.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="VendedorCod" HeaderText="Vendedor Cod.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Vendedor" HeaderText="Vendedor">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgLiberadoBloqueado" runat="server" Height="22px" ImageUrl="~/images/certo.jpg"
                                                data-ToolTip="default" ToolTip="Liberado" Width="22px" Style="border: 0;" />
                                            <asp:HiddenField ID="hidPedidoLiberado" runat="server" Value='<%# eval("Situacao") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPendente" ID="TabPendente">
                    <HeaderTemplate>
                        Pendente de Integração
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="titulodiv">
                            Pedidos pendentes de Integração
                        </div>
                        <div class="row">
                            <div class="bordagrid" style="height: 640px;">
                                <asp:GridView ID="gridPendente" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:BoundField DataField="VendedorCod" HeaderText="Vendedor Cod.">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="NomeVendedor" HeaderText="Nome">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PedidoNum" HeaderText="Número OnMobile">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ClienteCod" HeaderText="Cliente Cod.">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="NomeCliente" HeaderText="Nome">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="TabPrecoCod" HeaderText="Tabela">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PedidoData" HeaderText="Data" DataFormatString="{0:dd/MM/yyyy}">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PedidoVrpagar" HeaderText="Valor" DataFormatString="{0:n2}" HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
