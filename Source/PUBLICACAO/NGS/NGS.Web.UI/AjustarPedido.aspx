<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="AjustarPedido.aspx.vb" Inherits="NGS.Web.UI.AjustarPedido" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1360px !important;
        }

        .txtNumerico {
            text-align: right;
        }
    </style>
</asp:Content>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPedidosXItens" runat="server" AsyncPostBackTimeout="1000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPedidosXItens" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Ajustar Pedido
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkConfirmar" Text="Confirmar" runat="server" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="630px" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="630px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="591px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultarCliente" runat="server" Text=" > " OnClick="cmdConsultarCliente_Click" UseSubmitBehavior="False" CssClass="btn"
                        data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" Style="text-align: right;" runat="server" class="txtNumerico"
                        Width="100px" data-ToolTip="default" ToolTip="Número da Pedido." />
                </div>
                <div class="collbl">
                    Total anterior:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAnterior" Enabled="false" CssClass="txtDecimal" runat="server" />
                </div>
                <div class="collbl">
                    Total atual:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAtual" Enabled="false" CssClass="txtDecimal" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlOperacao" runat="server" Width="350px" AutoPostBack="true" />
                </div>
                <div class="collbl">
                    Sub-Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSubOperacao" runat="server" Width="350px" AutoPostBack="true" />
                </div>
            </div>
            <ajaxToolkit:TabContainer ID="tcPedido" runat="server" ActiveTabIndex="0" Width="100%"
                Style="clear: both;">
                <ajaxToolkit:TabPanel runat="server" HeaderText="Ites" ID="TabItens">
                    <HeaderTemplate>
                        Produto(s)
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="painelleft" style="width: 82%;">
                            <div class="bordagrid">
                                <asp:GridView ID="grdItens" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto">
                                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundField>
                                        <asp:TemplateField>
                                            <HeaderTemplate>
                                                Quantidade
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtQtd" Text='<%# String.Format("{0:N4}",eval("QuantidadePedidoFaturamento")) %>' Enabled="false" CssClass="txtDecimal4" runat="server" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <HeaderTemplate>
                                                Unitario
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtUnitario" Text='<%# String.Format("{0:N10}",eval("UnitarioMedioFaturamento")) %>' Enabled="false" CssClass="txtDecimal10" runat="server" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Lancamentos.TotalOficialPrd" HeaderText="Valor" DataFormatString="{0:n2}">
                                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundField>
                                        <asp:TemplateField ShowHeader="False">
                                            <ItemTemplate>
                                                <asp:Button ID="btnAlterarItem" runat="server" OnClick="btnAlterarItem_Click" Style="width: auto !important;"
                                                    Text="Alterar" />
                                                <asp:Button ID="btnSalvarItem" runat="server" CausesValidation="True" Text="OK" Visible="False" OnClick="btnSalvarItem_Click" Style="width: auto !important;" />
                                                <asp:Button ID="btnCancelarItem" runat="server" Text="Cancelar" Visible="False" Style="width: auto !important;"
                                                    OnClick="btnCancelarItem_Click" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" Width="100px" />
                                            <ItemStyle HorizontalAlign="Center" Width="100px" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Encargos">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lnkEncargos" runat="server" OnClick="lnkEncargos_Click"><img 
                                                src="Images/ico-mais.gif" alt="Encargos do Produto" /> </asp:LinkButton>
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <HeaderTemplate>
                                                Complemento R$
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtValorComplemento" Width="90px" CssClass="txtDecimal2" runat="server" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Left" Width="90px" />
                                            <ItemStyle HorizontalAlign="Left" Width="90px" />
                                        </asp:TemplateField>
                                        <asp:TemplateField ShowHeader="False">
                                            <ItemTemplate>
                                                <asp:Button ID="btnAlterarComplemento" runat="server" OnClick="btnAlterarComplemento_Click" Style="width: auto !important;"
                                                    Text="OK" data-ToolTip="default" ToolTip="Confirmar Complemento." />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                            <ItemStyle HorizontalAlign="Center" Width="30px" />
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <HeaderTemplate>
                                                Estorno R$
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtValorEstorno" Width="90px" CssClass="txtDecimal2" runat="server" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Left" Width="90px" />
                                            <ItemStyle HorizontalAlign="Left" Width="90px" />
                                        </asp:TemplateField>
                                        <asp:TemplateField ShowHeader="False">
                                            <ItemTemplate>
                                                <asp:Button ID="btnAlterarEstorno" runat="server" OnClick="btnAlterarEstorno_Click" Style="width: auto !important;"
                                                    Text="OK" data-ToolTip="default" ToolTip="Confirmar Estorno. Não mexe no Financeiro" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                            <ItemStyle HorizontalAlign="Center" Width="30px" />
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgExcluir" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                    OnClick="imgExcluir_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente excluir o produto?');" />
                                            </ItemTemplate>
                                            <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                            <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                        <div class="painelright" style="width: 17%;">
                            <div class="subtitulodiv">
                                Encargos Gerais do Pedido
                            </div>
                            <div class="bordagrid" style="height: 195px;">
                                <asp:GridView ID="gridEncargosGerais" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:BoundField DataField="CodigoEncargo" HeaderText="Encargo" />
                                        <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Reais">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ValorMoeda" DataFormatString="{0:N2}" HeaderText="Dolar">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                    </Columns>
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                    <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                    <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                    <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                </asp:GridView>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Vencimentos" ID="TabVencimentos">
                    <HeaderTemplate>
                        Vencimento(s)
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid" style="height: 225px;">
                            <asp:GridView ID="grdTitulos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None"
                                Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:BoundField DataField="TipoFinanceiro" HeaderText="Pagar/Receber">
                                        <HeaderStyle HorizontalAlign="Center" Width="100px" />
                                        <ItemStyle HorizontalAlign="Center" Width="100px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Titulo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Provisao" HeaderText="Provisao">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Vencimento" HeaderText="Vencimento">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescMoeda" HeaderText="Moeda"></asp:BoundField>
                                    <asp:BoundField DataField="DocumentoOficial" DataFormatString="{0:n2}" HeaderText="Doc. Oficial"></asp:BoundField>
                                    <asp:BoundField DataField="LiquidoOficial" DataFormatString="{0:N2}" HeaderText="Liq. Oficial">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DocumentoMoeda" DataFormatString="{0:N2}" HeaderText="Doc. Moeda">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="LiquidoMoeda" DataFormatString="{0:N2}" HeaderText="Liq. Moeda">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:PedidoXEncargo ID="ucPedidoxEncargo" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
