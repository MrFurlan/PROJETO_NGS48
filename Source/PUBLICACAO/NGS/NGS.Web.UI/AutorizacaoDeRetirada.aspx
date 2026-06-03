<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AutorizacaoDeRetirada.aspx.vb" Inherits="NGS.Web.UI.AutorizacaoDeRetirada" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 150px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngAutorizacaoDeRetirada" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAutorizacaoDeRetirada" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Autorização de Retirada
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
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
            <ajaxToolkit:TabContainer ID="tabContainer" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="tabCadastro" runat="server">
                    <HeaderTemplate>
                        Cadastro
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoEmpresa" runat="server" OnValueChanged="txtCodigoEmpresa_ValueChanged" />
                                <asp:TextBox ID="txtEmpresa" runat="server" Width="570px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnEmpresa" OnClick="btnEmpresa_Click" CssClass="btn" runat="server"
                                    Text=" > " UseSubmitBehavior="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                                <asp:TextBox ID="txtCliente" runat="server" Width="570px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnCliente" OnClick="btnCliente_Click" runat="server" Text=" > "
                                    CssClass="btn" UseSubmitBehavior="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedido" runat="server" Enabled="False" Font-Underline="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnPedido" OnClick="btnPedido_Click" runat="server" CssClass="btn"
                                    Text=">" UseSubmitBehavior="False" />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgExtratoPedido" OnClick="imgExtratoPedido_Click" runat="server"
                                    ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                    data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido" Visible="false" />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblOperacaoPedido" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblProdutoPedido" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Autorização:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAutorizacao" runat="server" Enabled="False" BackColor="White" />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgNfs" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                    OnClick="imgNfs_Click" data-ToolTip="default" ToolTip="Notas Vinculadas" Style="margin-top: 5px;" />
                            </div>
                            <div class="collbl">
                                Movimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMovimento" Width="70px" runat="server" CssClass="calendario" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Taxa:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTaxa" runat="server" AutoPostBack="True" CssClass="txtDecimal"
                                    OnTextChanged="txtTaxa_TextChanged" />
                                <asp:Label ID="lblTaxa" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Quantidade Fiscal:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtMaxQuantidadeFiscal" runat="server"></asp:HiddenField>
                                <asp:TextBox ID="txtQuantidadeFiscal" runat="server" AutoPostBack="True" CssClass="txtInteiro"
                                    OnTextChanged="txtQuantidadeFiscal_TextChanged" />
                            </div>
                            <div class="collbl" style="margin-left: 20px;">
                                Quantidade Física:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtMaxQuantidadeFisica" runat="server" />
                                <asp:TextBox ID="txtQuantidadeFisica" runat="server" AutoPostBack="True" CssClass="txtInteiro"
                                    OnTextChanged="txtQuantidadeFisica_TextChanged" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Contratado Fiscal:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQtdeContratada" runat="server" Enabled="False" CssClass="txtInteiro" />
                            </div>
                            <div class="collbl" style="margin-left: 20px;">
                                Contratado Físico:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQtdeContratadaFisica" runat="server" Enabled="False" CssClass="txtInteiro" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Entregue Fiscal:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQtdeEntregue" runat="server" Enabled="False" CssClass="txtInteiro" />
                            </div>
                            <div class="collbl" style="margin-left: 20px;">
                                Entregue Físico:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQtdeEntregueFisica" runat="server" Enabled="False" CssClass="txtInteiro" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Saldo Fiscal:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSaldoFiscal" runat="server" Enabled="False" CssClass="txtInteiro" />
                            </div>
                            <div class="collbl" style="margin-left: 20px;">
                                Saldo Físico:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSaldoFisico" runat="server" Enabled="False" CssClass="txtInteiro" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Autorizado Fiscal:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAutorizadoFiscal" runat="server" Enabled="False" CssClass="txtInteiro" />
                            </div>
                            <div class="collbl" style="margin-left: 20px;">
                                Autorizado Físico:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAutorizadoFisico" runat="server" Enabled="False" CssClass="txtInteiro" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente Retirada/Destino:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoClienteRetirada" runat="server" OnValueChanged="txtCodigoClienteRetirada_ValueChanged" />
                                <asp:TextBox ID="txtClienteRetirada" runat="server" Width="570px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnClienteRetirada" OnClick="btnClienteRetirada_Click" runat="server"
                                    CssClass="btn" Text=" > " UseSubmitBehavior="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Autorizante:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoAutorizante" runat="server" OnValueChanged="txtCodigoAutorizante_ValueChanged" />
                                <asp:TextBox ID="txtAutorizante" runat="server" Width="570px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnClienteAutorizante" OnClick="btnClienteAutorizante_Click" runat="server"
                                    CssClass="btn" Text=" > " UseSubmitBehavior="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data Base de Cálculo:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataBase" runat="server" Width="86px" CssClass="calendario" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido Serviço:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedidoServico" runat="server" Width="86px" Enabled="false" CssClass="txtNumerico" />
                            </div>
                            <div class="collbl" style="width: 113px;">
                                Título:
                            </div>
                            <div class="coltxt" style="text-align: right">
                                <asp:TextBox ID="txtTitulo" runat="server" Width="64px" Enabled="False" CssClass="txtNumerico" />
                            </div>
                            <div class="collbl" style="width: 113px;">
                                Venc. Título:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVencimento" runat="server" Width="70px" CssClass="calendario" />
                            </div>
                            <div class="collbl" style="width: 113px;">
                                Valor:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtValor" runat="server" Width="86px" Enabled="False" CssClass="txtDecimal" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observação:
                            </div>
                            <div class="coltxt" style="width: 78.8%;">
                                <asp:TextBox ID="txtObservacao" runat="server" Width="100%" OnTextChanged="txtObservacao_TextChanged"
                                    TextMode="MultiLine" />
                            </div>
                        </div>
                        <div class="coltxt" style="margin-left: 161px;">
                            <div class="coltxt">
                                <asp:Button ID="BtnObservacao" runat="server" CssClass="botao" Text="Observação" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="tabConsulta" runat="server">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="gridAut" runat="server" AutoGenerateColumns="False" OnSelectedIndexChanged="gridAut_SelectedIndexChanged"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Font-Size="Smaller" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:BoundField DataField="Autorizacao" HeaderText="ID" />
                                    <asp:BoundField DataField="CodigoPedido" HeaderText="Pedido" />
                                    <asp:BoundField DataField="NomeClientePedido" HeaderText="Cliente" />
                                    <asp:BoundField DataField="NomeClienteRetirada" HeaderText="Destinatário">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeAutorizante" HeaderText="Autorizante" DataFormatString="{0:N4}"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Taxa" HeaderText="Taxa" DataFormatString="{0:N6}" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="QuantidadeAutorizadaFisica" HeaderText="Autor.F&#237;sico"
                                        DataFormatString="{0:N0}" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="QuantidadeEntregueFisica" HeaderText="Entregue F&#237;sico"
                                        DataFormatString="{0:N0}" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="SaldoFisico" HeaderText="Saldo F&#237;sico" DataFormatString="{0:N0}"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="QuantidadeAutorizadaFiscal" DataFormatString="{0:N0}"
                                        HeaderText="Autor.Fiscal" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="QuantidadeEntregueFiscal" DataFormatString="{0:N0}" HeaderText="Entregue Fiscal"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="SaldoFiscal" DataFormatString="{0:N0}" HeaderText="Saldo Fiscal"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoPedidoServico" HeaderText="Pedido Servico">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>

        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:ConsultaAutorizacaoDeRetirada ID="ucConsultaAutorizacaoDeRetirada" runat="server" />
    <uc:ConsultaObservacoes ID="ucConsultaObservacoes" runat="server" />
</asp:Content>
