<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="LiberacaoDePedido.aspx.vb" Inherits="NGS.Web.UI.LiberacaoDePedido" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }
        .collbl {
            width: 140px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngLiberacaoDePedido" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlLiberacaoDePedido" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Liberação de Pedido
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
                                <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente."
                                    runat="server" Text="Consolidar Cliente:" />
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
                                Usuário Liberação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUsuarioLiberacao" runat="server" Width="650px" />
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
                                &nbsp;&nbsp;
                                <asp:CheckBox ID="chkPeriodo" data-ToolTip="default" ToolTip="Considerar período." Text="Considerar período:" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtConsultaPedido" runat="server" Width="100px" CssClass="txtNumerico" data-ToolTip="default" ToolTip="Número do pedido a ser consultado." />
                            </div>
                            <div class="collbl">
                                Tipo:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdBloqueados" runat="server" Checked="True" GroupName="liberacao"
                                    Text="Bloqueados" data-ToolTip="default" ToolTip="Selecionar a opção paravisualizar pedidos bloqueados." />
                                <asp:RadioButton ID="rdLiberados" runat="server" GroupName="liberacao" Text="Liberados" data-ToolTip="default" ToolTip="Selecionar a opção paravisualizar pedidos liberados." />
                                <asp:RadioButton ID="rdTodos" runat="server" GroupName="liberacao" Text="Todos" data-ToolTip="default" ToolTip="Selecionar a opção paravisualizar pedidos bloqueados e liberados." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Troca:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdTrocaSim" runat="server" GroupName="Troca" Text="Sim" data-ToolTip="default" ToolTip="Marcar se o pedido refere-se ou não a uma troca." />
                                <asp:RadioButton ID="rdTrocaNao" runat="server" GroupName="Troca" Text="Não" data-ToolTip="default" ToolTip="Marcar se o pedido refere-se ou não a uma troca." />
                                <asp:RadioButton ID="rdTrocaTodos" runat="server" GroupName="Troca" Text="Todos"
                                    Checked="true" data-ToolTip="default" ToolTip="Marcar se o pedido refere-se ou não a uma troca." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridConsulta" runat="server" CellPadding="4" ForeColor="#333333"
                                GridLines="None" Width="99%" AutoGenerateColumns="False">
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
                                    <asp:BoundField DataField="Classe" HeaderText="Opera&#231;&#227;o">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Produto" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" DataFormatString="{0:N4}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Unitario" HeaderText="Unit&#225;rio" DataFormatString="{0:N10}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" HeaderText="Valor" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Moeda">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="UsuarioLiberacao" HeaderText="Liberação">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="UsuarioLiberacaoData" HeaderText="Data" DataFormatString="{0:dd/MM/yyyy}">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgStatus" runat="server" Height="20px" Width="20px" ImageAlign="AbsMiddle"
                                                ImageUrl="~/images/certo.jpg" Style="border: 0;" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
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
                        <div class="painelleft" style="width: 33%; margin-right: 4px;">
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
                        <div class="painelleft" style="width: 33%; margin-right: 4px;">
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
                        <div class="painelleft" style="width: 33%;">
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
                        </div>
                        <div class="painelleft" style="width: 33%; margin-right: 4px;">
                            <div class="subtitulodiv">
                                PRODUTO
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    NOME:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtProduto" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    QUANTIDADE:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtQuantidade" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    UNITÁRIO R$:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtUnitario" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    VALOR R$:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtValor" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    UNITÁRIO U$:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtUnitarioMoeda" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    VALOR U$:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtValorMoeda" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    ÍNDICE FIXADO:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtIndiceFixado" runat="server" />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 66%;">
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
                        <div style="clear: both;" runat="server">
                            <div class="subtitulodiv">
                                ENCARGOS
                            </div>
                            <div class="bordagrid" style="height: 120px;">
                                <asp:GridView ID="gridEncargos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <Columns>
                                        <asp:BoundField DataField="Codigo" HeaderText="Encargo">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="SituacaoTributaria" HeaderText="ClaICMS">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="SituacaoTributariaPISCOFINS" HeaderText="PisCofins">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="SituacaoTributariaIPI" HeaderText="IPI">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Base" DataFormatString="{0:N2}" HeaderText="Base">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Percentual" DataFormatString="{0:N6}" HeaderText="% Calc.">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PercentualExibicao" DataFormatString="{0:N6}" HeaderText="% Exib.">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Valor" DataFormatString="{0:N2}" HeaderText="Valor">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Sinal" HeaderText="Sinal">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                    </Columns>
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <EditRowStyle BackColor="#999999" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                </asp:GridView>
                            </div>
                        </div>
                        <div class="subtitulodiv" style="clear: both;">
                            FRETE: &nbsp;<asp:Label ID="txtCifFob" runat="server" ForeColor="Yellow" />
                        </div>
                        <div runat="server">
                            <div class="bordagrid" style="height: 120px;">
                                <asp:GridView ID="gridTransportes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="99%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <Columns>
                                        <asp:BoundField DataField="Codigo" HeaderText="CNPJ">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CodigoEndereco" HeaderText="End">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Nome">
                                            <ItemTemplate>
                                                <asp:Label ID="Label1" runat="server" Text='<%# Bind("Transportador.Nome") %>' />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="DataFrete" HeaderText="Data">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="UnitarioFrete" HeaderText="Unitário">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                    </Columns>
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <EditRowStyle BackColor="#999999" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                </asp:GridView>
                            </div>
                        </div>
                        <div runat="server">
                            <div class="subtitulodiv">
                                REPRESENTANTE
                            </div>
                            <div class="bordagrid" style="height: 120px;">
                                <asp:GridView ID="gridRepresentantes" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="99%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <Columns>
                                        <asp:BoundField DataField="CodigoRepresentante" HeaderText="Codigo">
                                            <HeaderStyle HorizontalAlign="Left" Width="110px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left" Width="110px"></ItemStyle>
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CodigoEnderecoRepresentante" HeaderText="End.">
                                            <HeaderStyle HorizontalAlign="Left" Width="10px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left" Width="10px"></ItemStyle>
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Nome">
                                            <ItemTemplate>
                                                <asp:Label ID="Label1" runat="server" Text='<%# Bind("Representante.Nome") %>' />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Percentual" DataFormatString="{0:n2}" HeaderText="Percentual">
                                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ValorComissao" DataFormatString="{0:n2}" HeaderText="Valor">
                                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundField>
                                    </Columns>
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <EditRowStyle BackColor="#999999" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                </asp:GridView>
                            </div>
                        </div>
                        <div style="clear: both;">
                            <div class="subtitulodiv">
                                FINANCEIRO
                            </div>
                            <div class="row">
                                <div class="collbl" style="width: 125px;">
                                    Momento Financeiro:
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="ajustarFinanceiro" runat="server" />
                                    <asp:DropDownList ID="ddlMomentoFinanceiro" runat="server" AutoPostBack="True" Width="240px"
                                        Enabled="False">
                                        <asp:ListItem Value="0">Nenhum</asp:ListItem>
                                        <asp:ListItem Value="2">Venc. No Pedido</asp:ListItem>
                                        <asp:ListItem Value="3">Venc. Na Nota</asp:ListItem>
                                        <asp:ListItem Value="4">Troca</asp:ListItem>
                                        <asp:ListItem Value="9">Virtual</asp:ListItem>
                                    </asp:DropDownList>
                                </div>
                                <div class="collbl">
                                    Pagamento:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="lstCondicoes" runat="server" Width="420px" AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl" style="width: 125px;">
                                    Data:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataVencimento" CssClass="calendario" runat="server" Width="100px" />
                                </div>
                                <div class="collbl" style="margin-left: 130px;">
                                    Valor:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtValorVencimento" CssClass="txtDecimal" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnOkVencimento" CssClass="botao" runat="server" Text="OK" UseSubmitBehavior="False" />
                                </div>
                            </div>
                            <div class="bordagrid" style="height: 120px;">
                                <asp:GridView ID="gridFinanceiro" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridFinanceiro_SelectedIndexChanged"
                                    Width="100%">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:TemplateField ShowHeader="False">
                                            <ItemTemplate>
                                                <asp:Button ID="Button1" runat="server" CausesValidation="False" CommandName="Select"
                                                    Enabled='<%# Eval("Provisao") = 2 %>' Text="&gt;" />
                                            </ItemTemplate>
                                            <ControlStyle Width="30px" />
                                            <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                            <ItemStyle HorizontalAlign="Center" Width="30px" />
                                        </asp:TemplateField>
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
                                        <asp:BoundField DataField="DescMoeda" HeaderText="Moeda"></asp:BoundField>
                                        <asp:BoundField DataField="DocumentoOficial" DataFormatString="{0:n2}" HeaderText="Doc. Oficial"></asp:BoundField>
                                        <asp:BoundField DataField="AcrescimoOficial" DataFormatString="{0:N2}" HeaderText="Acres. Oficial">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="DeducaoOficial" DataFormatString="{0:N2}" HeaderText="Dedução Oficial">
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
                                        <asp:BoundField DataField="AcrescimoMoeda" DataFormatString="{0:N2}" HeaderText="Acres. Moeda">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="DeducaoMoeda" DataFormatString="{0:N2}" HeaderText="Deduções Moeda">
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
                        <asp:Panel ID="pnlEntrega" runat="server" Width="99%" Visible="true">
                            <div class="row">
                                <div class="collbl" style="width: 230px;">
                                    CONDIÇÃO DE PAGAMENTO DA ENTREGA:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="lstCondicoesPgtoEntrega" runat="server" Width="420px" OnSelectedIndexChanged="lstCondicoesPgtoEntrega_SelectedIndexChanged"
                                        AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl" style="width: 230px;">
                                    QUOTA DA ENTREGA
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtQuotaDeEntrega" CssClass="txtDecimal4" runat="server" Text="0"
                                        AutoPostBack="True" OnTextChanged="txtQuotaDeEntrega_TextChanged" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl" style="width: 230px;">
                                    PERIODICIDADE DA ENTREGA
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlPeriodicidadeEntrega" runat="server" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlPeriodicidadeEntrega_SelectedIndexChanged" Width="420px">
                                        <asp:ListItem Value="0">Antecipado/Fabrica/Avista</asp:ListItem>
                                        <asp:ListItem Value="1">Na Entrega</asp:ListItem>
                                        <asp:ListItem Value="2">Data Fixa</asp:ListItem>
                                        <asp:ListItem Value="7">Semanal Entrega de segunda a domingo com Pagamento X Dias após</asp:ListItem>
                                        <asp:ListItem Value="15">Quinzenal Entrega durante a quinzena com pagamento X Dias após</asp:ListItem>
                                    </asp:DropDownList>
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnAjustarEntrega" runat="server" Text="Ajustar" CssClass="botao" />
                                </div>
                            </div>
                        </asp:Panel>
                        <hr />
                        <div class="row">
                            <div class="coltxt">
                                <asp:Button ID="btnLiberar" runat="server" Text="Liberar" BackColor="Green" CssClass="btn"
                                    ForeColor="White" Style="cursor: pointer; font-family: Tahoma,Arial,Helvetica,sans-serif; font-size: 11px; height: 24px; text-align: center; width: 96px;"
                                    OnClick="btnLiberar_Click" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
