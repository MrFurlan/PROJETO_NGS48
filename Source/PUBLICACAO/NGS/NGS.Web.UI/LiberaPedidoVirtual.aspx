<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="LiberaPedidoVirtual.aspx.vb" Inherits="NGS.Web.UI.LiberaPedidoVirtual" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1360px !important;
        }
    </style>

    <script type="text/javascript">
        function downloadArquivo() {
            alert("Virtualizado com sucesso");
            $("#cmdArquivoDeSaida").click();
        }
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngLiberaPedidoVirtual" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlLiberaPedidoVirtual" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="HiddenArquivo" runat="server" />
            <div class="titulodiv">
                Liberação de Pedido Virtual
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Consultar
                    </HeaderTemplate>
                    <ContentTemplate>
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
                                <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="650px" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
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
                                <asp:TextBox ID="txtCliente" runat="server" Width="611px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnCliente" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False"
                                    OnClick="btnCliente_Click" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Período:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" data-ToolTip="default"
                                    ToolTip="Informar o período a ser considerado." Width="100px" />
                                &nbsp; a&nbsp;
                                <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" data-ToolTip="default"
                                    ToolTip="Informar o período a ser considerado." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigoPedido" runat="server" CssClass="txtNumerico9" Width="100px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Moeda:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdReal" runat="server" GroupName="Moeda" Text="Real" Checked="True" data-ToolTip="default" ToolTip="Real" />
                                <asp:RadioButton ID="rdDolar" runat="server" GroupName="Moeda" Text="Dólar" data-ToolTip="default" ToolTip="Dólar" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                E/S:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdEntrada" runat="server" GroupName="Financeiro" Text="Entrada" Checked="True" data-ToolTip="default" ToolTip="Pedidos de Entrada." />
                                <asp:RadioButton ID="rdSaida" runat="server" GroupName="Financeiro" Text="Saída" data-ToolTip="default" ToolTip="Pedidos de Saída." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tipo:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdAberto" runat="server" GroupName="Tipo" Text="Abertos" Checked="True" data-ToolTip="default" ToolTip="Apenas título(s) em aberto." />
                                <asp:RadioButton ID="rdAmbos" runat="server" GroupName="Tipo" Text="Ambos" data-ToolTip="default" ToolTip="Título(s) aberto(s) e baixado(s)." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridConsulta" runat="server" CellPadding="4" ForeColor="#333333"
                                GridLines="None" Width="100%" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:LinkButton ID="lnkSelecionar" CssClass="lnk" 
                                                data-tooltip="default" ToolTip="Selecionar registro." runat="server" Text=" &gt; "
                                                OnClick="lnkSelecionar_Click">
                                                <i class="fa fa-arrow-right seta"></i>
                                            </asp:LinkButton>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeCliente" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Pedido" HeaderText="Pedido">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="PedidoEfetivo" HeaderText="CN">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Data" HeaderText="Data" DataFormatString="{0:dd/MM/yyyy}">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Moeda">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" DataFormatString="{0:N4}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorPedido" HeaderText="Valor Pedido" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="valoraberto" HeaderText="Finan.Aberto" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="valorbaixado" HeaderText="Finan.Baixado" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
                    <HeaderTemplate>
                        Liberar
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkUmNovo" Text="Gravar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 49%; margin-right: 4px;">
                            <div class="subtitulodiv">
                                EMPRESA
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    CNPJ:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtCnpjDaEmpresa" runat="server" />
                                    &nbsp;
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    NOME/RAZÃO:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtNomeDaEmpresa" runat="server" Font-Bold="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    CIDADE/UF:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtCidadeDaEmpresa" runat="server" />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 49%">
                            <div class="subtitulodiv">
                                CLIENTE/FORNECEDOR
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    CPF/CNPJ:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtCnpjCliente" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    NOME/RAZÃO:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtNomeDoCliente" runat="server" Font-Bold="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    CIDADE/UF:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtCidadeCliente" runat="server" />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 49%; margin-right: 4px;">
                            <div class="subtitulodiv">
                                DEPÓSITO
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    CNPJ:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtDeposito" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    NOME/RAZÃO:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtNomeDoDeposito" runat="server" Font-Bold="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    CIDADE/UF:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtCidadeDeposito" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div>
                                    &nbsp;
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Total Titulos:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtTotalTitulos" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal" Font-Bold="True"
                                        ForeColor="Red" Style="width: 91px;" data-ToolTip="default" ToolTip="Valor total do(s) Título(s)." ReadOnly="True" />
                                </div>
                                <div class="collbl" style="margin-left: 30px;">
                                    Total Notas:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtTotalNotas" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal"
                                        Font-Bold="True" ForeColor="Red" Style="width: 91px;" data-ToolTip="default" ToolTip="Valor total da(s) Nota(s)." ReadOnly="True" />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 49%;">
                            <div class="subtitulodiv">
                                COMERCIALIZAÇÃO
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    PEDIDO:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtPedido" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgExtratoPedido" runat="server" CssClass="imgconsultar" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                        OnClick="imgExtratoPedido_Click" data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido" />
                                </div>
                                <div>
                                    <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                                        <ContentTemplate>
                                            <asp:Button ID="cmdArquivoDeSaida" CssClass="botao" runat="server" Height="20px"
                                                Text="Download" UseSubmitBehavior="False" ClientIDMode="Static" Visible="false" />
                                        </ContentTemplate>
                                        <Triggers>
                                            <asp:PostBackTrigger ControlID="cmdArquivoDeSaida" />
                                        </Triggers>
                                    </asp:UpdatePanel>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    ENTREGA:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataEntrega" runat="server" Width="85px" BorderStyle="None" AutoPostBack="True"
                                        ClientIDMode="Static" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    MOEDA:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtMoeda" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    NAT. OPERAÇÃO:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtNaturezaDaOperacao" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    OP. COMERCIAL:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtOperacao" runat="server" />
                                </div>
                            </div>
                        </div>
                        <div style="clear: both;">
                            <div class="subtitulodiv">
                                FINANCEIRO
                            </div>
                            <div class="bordagrid" style="height: 120px;">
                                <asp:GridView ID="gridFinanceiro" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
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
                                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                            <ItemStyle HorizontalAlign="Left" Width="100px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="DataBaixa" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Pagamento"
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Moeda" HeaderText="Moeda"></asp:BoundField>
                                        <asp:BoundField DataField="DocumentoOficial" DataFormatString="{0:n2}" HeaderText="Doc. Oficial">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="DeducaoOficial" DataFormatString="{0:N2}" HeaderText="Dedução Oficial">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="DescontoOficial" DataFormatString="{0:N2}" HeaderText="Desconto Oficial">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="AcrescimoOficial" DataFormatString="{0:N2}" HeaderText="Acres. Oficial">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="JurosOficial" DataFormatString="{0:N2}" HeaderText="Juros Oficial">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="LiquidoOficial" DataFormatString="{0:N2}" HeaderText="Liq. Oficial">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="DocumentoMoeda" DataFormatString="{0:N2}" HeaderText="Doc. Moeda">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="DeducaoMoeda" DataFormatString="{0:N2}" HeaderText="Deduções Moeda">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="DescontoMoeda" DataFormatString="{0:N2}" HeaderText="Deduções Moeda">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="AcrescimoMoeda" DataFormatString="{0:N2}" HeaderText="Acres. Moeda">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="JurosMoeda" DataFormatString="{0:N2}" HeaderText="Juros Moeda">
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
                        </div>
                        <div style="clear: both;">
                            <div class="subtitulodiv">
                                NOTAS FISCAIS
                            </div>
                            <div class="bordagrid" style="height: 120px;">
                                <asp:GridView ID="GridNotas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:BoundField DataField="Movimento" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Empresa_id" HeaderText="Empresa">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="EndEmpresa_id" HeaderText="End">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Cliente" HeaderText="Cliente" HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="EndCliente" HeaderText="End">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="NomeCliente" HeaderText="Nome">
                                            <HeaderStyle HorizontalAlign="Left" Width="350px" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ES" HeaderText="E/S">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Serie" HeaderText="S&#233;rie">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Nota" HeaderText="Nota">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Operacao" HeaderText="OP">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="SubOperacao" HeaderText="SO">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Valor" HeaderText="Valor" DataFormatString="{0:n2}">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabParcelamento" runat="server" HeaderText="TabPanel3">
                    <HeaderTemplate>
                        Financeiro Virtual
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconMais" runat="server">
                                        <asp:LinkButton ID="LnkParcelar" Text="Parcelar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkVirtualizar" Text="Gerar Virtualização" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkVincular" Text="Gerar Vínculos" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="painelleft" style="width: 40%;">
                                <div class="row">
                                    <div class="collbl" style="width: 100px;">
                                        Condições:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlCondPagPed" runat="server" AutoPostBack="True" Width="357px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl" style="width: 100px;">
                                        Data:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtDataCondPagParcela" runat="server" CssClass="calendario" Width="100px" />
                                    </div>
                                    <div class="collbl" style="width: 100px;">
                                        Total Pedido:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtPedidoTotal" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal" Font-Bold="True" ForeColor="Red" Text="0" data-ToolTip="default" ToolTip="Valor que será parcelado." ReadOnly="True" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl" style="width: 100px;">
                                        Total Realizado:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtPedidoTotalPago" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal" Font-Bold="True" ForeColor="Red" Text="0" data-ToolTip="default" ToolTip="Valor das parcelas pagas." ReadOnly="True" />
                                    </div>
                                    <div class="collbl" style="width: 100px; margin-left: 21px;">
                                        Saldo:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtPedidoSaldo" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal" Font-Bold="True" ForeColor="Red" Text="0" data-ToolTip="default" ToolTip="Saldo a pagar." ReadOnly="True" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="painelleft" style="width: 45%;">
                                        <div class="bordagrid" style="height: 215px;">
                                            <asp:GridView ID="gridParcelas" runat="server" CellPadding="4"
                                                ForeColor="#333333" GridLines="None"
                                                Width="100%" AutoGenerateColumns="False">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                                    <asp:BoundField DataField="CodigoParcela" HeaderText="Parcela">
                                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="DataVencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Valor" DataFormatString="{0:N2}" HeaderText="Valor">
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
                                    </div>
                                    <div class="painelright" style="width: 54%">
                                        <div class="bordagrid" style="height: 215px;">
                                            <div class="menu_acoes" style="width: 96%;">
                                                <div class="acoes">
                                                    <ul>
                                                        <li class="icon_confirmar" runat="server">
                                                            <asp:LinkButton class="iconAtualizar" ID="LnkAtualizarParcela" Text="Atualizar" runat="server" />
                                                        </li>
                                                    </ul>
                                                </div>
                                            </div>
                                            <div class="row" style="margin-left: 1px;">
                                                <div class="collbl" style="width: 100px;">
                                                    Parcela:
                                                </div>
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtCodigoParcela" runat="server" CssClass="inteiro" Width="100px" BackColor="#FFFFC0" ReadOnly="True" />
                                                </div>
                                            </div>
                                            <div class="row" style="margin-left: 1px;">
                                                <div class="collbl" style="width: 100px;">
                                                    Data:
                                                </div>
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtDataVencParcela" runat="server" CssClass="calendario" Width="100px" />
                                                </div>
                                            </div>
                                            <div class="row" style="margin-left: 1px;">
                                                <div class="collbl" style="width: 100px;">
                                                    Valor:
                                                </div>
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtValorParcela" runat="server" CssClass="txtDecimal" Width="100px" />
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="painelright" style="width: 59%;">
                                <div class="bordagrid">
                                    <asp:GridView ID="GridNotasXTitulos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                        ForeColor="#333333" GridLines="None" Width="100%">
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <EditRowStyle BackColor="#999999" />
                                        <Columns>
                                            <asp:BoundField DataField="ES" HeaderText="E/S">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Serie" HeaderText="S&#233;rie">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Nota" HeaderText="Nota">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TipoFinanceiro" HeaderText="Pag/Rec">
                                                <HeaderStyle HorizontalAlign="Center" Width="100px" />
                                                <ItemStyle HorizontalAlign="Center" Width="100px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Provisao" HeaderText="Provisão">
                                                <HeaderStyle HorizontalAlign="Center" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Titulo" HeaderText="Título">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Venc">
                                                <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor R$">
                                                <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="DeducaoOficial" DataFormatString="{0:N2}" HeaderText="Dedução R$">
                                                <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="DescontoOficial" DataFormatString="{0:N2}" HeaderText="Desconto R$">
                                                <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="AcrescimoOficial" DataFormatString="{0:N2}" HeaderText="Acréscimo R$">
                                                <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="JurosOficial" DataFormatString="{0:N2}" HeaderText="Juros R$">
                                                <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="LiquidoOficial" DataFormatString="{0:N2}" HeaderText="Líquido R$">
                                                <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="LiquidoMoeda" DataFormatString="{0:N2}" HeaderText="Líquido U$">
                                                <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
