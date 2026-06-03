<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CorrecaoMonetaria.aspx.vb" Inherits="NGS.Web.UI.CorrecaoMonetaria" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1180px !important;
        }
    </style>
    <script type="text/javascript">
        function selectAll(chkAll, grid) {
            var chk = $('#' + chkAll.id);
            var checked = chk.attr('checked') == "checked";

            $("input[type='checkbox']", grid).not(".chkAlls").each(function () {
                $(this).attr("checked", checked);
            });
        };
    </script>
    <style type="text/css">
        .collbl {
            width: 140px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCorrecaoMonetaria" runat="server" AsyncPostBackTimeout="1000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCorrecaoMonetaria" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv" style="line-height: 32px; margin-bottom: 5px;">
                <label>
                    Correção Monetária Pedidos
                </label>
            </div>
            <ajaxToolkit:TabContainer ID="TC01" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TPEmAndamento" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Pedidos Em Andamento
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server" style="width: 150px">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar Pedidos"
                                            runat="server" />
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
                                <asp:CheckBox ID="ChkConsEmpresa" runat="server" AutoPostBack="True" Text="Consolidar Empresa:" />
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Safra:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSafra" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlSafra_SelectedIndexChanged"
                                    Width="600px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                <asp:CheckBox ID="chkConsolidarCliente" data-ToolTip="default" ToolTip="Consolidar o cpf/cnpj do cliente."
                                    runat="server" AutoPostBack="True" OnCheckedChanged="chkConsolidarCliente_CheckedChanged"
                                    Text="Consolidar Cliente:" />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeCliente" Enabled="False" runat="server" Width="567px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnEmitente" runat="server" OnClick="btnEmitente_Click" Text=">"
                                    CssClass="btn" data-ToolTip="default" ToolTip="Unificar as informações por cliente." />
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Mês/Ano Correção:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlMes" runat="server">
                                    <asp:ListItem Value="01">Janeiro</asp:ListItem>
                                    <asp:ListItem Value="02">Fevereiro</asp:ListItem>
                                    <asp:ListItem Value="03">Marco</asp:ListItem>
                                    <asp:ListItem Value="04">Abril</asp:ListItem>
                                    <asp:ListItem Value="05">Maio</asp:ListItem>
                                    <asp:ListItem Value="06">Junho</asp:ListItem>
                                    <asp:ListItem Value="07">Julho</asp:ListItem>
                                    <asp:ListItem Value="08">Agosto</asp:ListItem>
                                    <asp:ListItem Value="09">Setembro</asp:ListItem>
                                    <asp:ListItem Value="10">Outubro</asp:ListItem>
                                    <asp:ListItem Value="11">Novembro</asp:ListItem>
                                    <asp:ListItem Value="12">Dezembro</asp:ListItem>
                                </asp:DropDownList>
                                <asp:DropDownList ID="ddlAno" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Indexador:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlIndexador" runat="server" OnSelectedIndexChanged="ddlIndexador_SelectedIndexChanged"
                                    AutoPostBack="True" Width="600px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Indice:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtIndice" runat="server" data-ToolTip="default" ToolTip="Cotação da moeda." />
                            </div>
                        </div>
                        <ajaxToolkit:TabContainer ID="TC02" runat="server" ActiveTabIndex="1" Width="100%">
                            <ajaxToolkit:TabPanel ID="TPPedidos" runat="server" HeaderText="TabPanel4">
                                <HeaderTemplate>
                                    Pedidos com Variação
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="menu_acoes">
                                        <div class="acoes">
                                            <ul>
                                                <li runat="server" style="width: 140px">
                                                    <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Visual. Correção" runat="server" />
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                    <div class="bordagrid">
                                        <asp:GridView ID="gridPedidos" runat="server" AutoGenerateColumns="False" CellPadding="4"
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
                                                        <input type="checkbox" id="chkAll" onclick="selectAll(chkAll, '#MainContent_TC01_TPEmAndamento_TC02_TPPedidos_gridPedidos');" />
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:CheckBox ID="chkPedido" runat="server" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="Empresa" HeaderText="Empresa" />
                                                <asp:BoundField DataField="EndEmpresa" />
                                                <asp:BoundField DataField="Fantasia" />
                                                <asp:BoundField DataField="Cliente" HeaderText="Cliente" />
                                                <asp:BoundField DataField="EndCliente" />
                                                <asp:BoundField DataField="Nome" HeaderText="Nome" />
                                                <asp:BoundField DataField="Pedido" HeaderText="Pedido" />
                                                <asp:BoundField DataField="Safra" HeaderText="Safra" />
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:ImageButton ID="imgExtratoPedido" runat="server" Height="23px" ImageAlign="AbsMiddle"
                                                            ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imgExtratoPedido_Click"
                                                            data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido" Width="23px" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TPVisualizacao" runat="server" HeaderText="TabPanel2">
                                <ContentTemplate>
                                    <div class="menu_acoes">
                                        <div class="acoes">
                                            <ul>
                                                <li runat="server" style="width: 140px;">
                                                    <asp:LinkButton class="iconRelatorio" ID="lnkRelatorioVC" Text="Visual.Contabilidade"
                                                        runat="server" />
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                    <div class="bordagrid">
                                        <asp:GridView ID="gridCorrecao" runat="server" CellPadding="4" ForeColor="#333333"
                                            GridLines="None" AutoGenerateColumns="False">
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
                                                        <input type="checkbox" id="chkTodos" onclick="selectAll(this, '#MainContent_TC01_TPEmAndamento_TC02_TPVisualizacao_gridCorrecao');" />
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:CheckBox ID="chkPedidoCorrecao" runat="server" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="DescTipo" HeaderText="Tipo" />
                                                <asp:BoundField DataField="Empresa" HeaderText="Empresa" />
                                                <asp:BoundField DataField="EndEmpresa" />
                                                <asp:BoundField DataField="Fantasia" />
                                                <asp:BoundField DataField="Cliente" HeaderText="Cliente" />
                                                <asp:BoundField DataField="EndCliente" />
                                                <asp:BoundField DataField="Nome" />
                                                <asp:BoundField DataField="Pedido" HeaderText="Pedido" />
                                                <asp:BoundField DataField="VariacaoCambial" HeaderText="Var. Cam. Pedido" />
                                                <asp:BoundField DataField="VariacaoPassivaCorrecao" HeaderText="Var. Razao Passiva" />
                                                <asp:BoundField DataField="VariacaoAtivaCorrecao" HeaderText="Var. Razao Ativa" />
                                                <asp:BoundField DataField="VariacaoCorrigida" HeaderText="Variacao Final" />
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:ImageButton ID="imgExtratoPedido2" runat="server" Height="23px" ImageAlign="AbsMiddle"
                                                            ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imgExtratoPedido2_Click"
                                                            data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido" Width="23px" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="FinanceiroBaixadoEmDolar" DataFormatString="{0:N2}" HeaderText="Fin.Bx.Dolar" />
                                                <asp:BoundField DataField="FinanceiroBaixadoEmReais" DataFormatString="{0:N2}" HeaderText="Fin.Bx.Reais" />
                                                <asp:BoundField DataField="ValorDolarPedido" DataFormatString="{0:N2}" HeaderText="Vlr Dolar Pedido" />
                                                <asp:BoundField DataField="ValorEmDolarPedidoAtualizadoEmReais" DataFormatString="{0:N2}"
                                                    HeaderText="Vlr Reais Pedido" />
                                                <asp:BoundField DataField="ReaisNotas" DataFormatString="{0:N2}" HeaderText="Reais NF Emitidas" />
                                                <asp:BoundField />
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </ContentTemplate>
                                <HeaderTemplate>
                                    Visualizar Correção
                                </HeaderTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TPContabilizacao" runat="server" HeaderText="TabPanel3">
                                <HeaderTemplate>
                                    Contabilização
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="menu_acoes">
                                        <div class="acoes">
                                            <ul>
                                                <li runat="server" style="width: 140px;">
                                                    <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Contabilizar" runat="server" />
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                    <div class="bordagrid">
                                        <asp:GridView ID="gridContabilizacao" runat="server" CellPadding="4" ForeColor="#333333"
                                            GridLines="None" AutoGenerateColumns="False" Width="100%">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EditRowStyle BackColor="#999999" />
                                            <Columns>
                                                <asp:BoundField DataField="Pedido" HeaderText="Pedido" />
                                                <asp:BoundField DataField="Conta" HeaderText="Conta" />
                                                <asp:BoundField DataField="Cliente" HeaderText="Cliente" />
                                                <asp:BoundField DataField="EndCliente" />
                                                <asp:BoundField DataField="Nome" HeaderText="Nome" />
                                                <asp:BoundField DataField="Historico" HeaderText="Historico" />
                                                <asp:BoundField DataField="DebitoOficial" HeaderText="Deb. Oficial" />
                                                <asp:BoundField DataField="creditooficial" HeaderText="Cred. Oficial" />
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                        </ajaxToolkit:TabContainer>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TPEncerrado" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Pedidos Encerrados Estorno da Provisão
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server" style="width: 200px">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultaEncerrados" Text="Consultar Pedidos Encerrados"
                                            runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaEncerrados" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Mês/Ano Correção:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlMesEncerrado" runat="server">
                                    <asp:ListItem Value="01">Janeiro</asp:ListItem>
                                    <asp:ListItem Value="02">Fevereiro</asp:ListItem>
                                    <asp:ListItem Value="03">Marco</asp:ListItem>
                                    <asp:ListItem Value="04">Abril</asp:ListItem>
                                    <asp:ListItem Value="05">Maio</asp:ListItem>
                                    <asp:ListItem Value="05">Junho</asp:ListItem>
                                    <asp:ListItem Value="07">Julho</asp:ListItem>
                                    <asp:ListItem Value="08">Agosto</asp:ListItem>
                                    <asp:ListItem Value="09">Setembro</asp:ListItem>
                                    <asp:ListItem Value="10">Outubro</asp:ListItem>
                                    <asp:ListItem Value="11">Novembro</asp:ListItem>
                                    <asp:ListItem Value="12">Dezembro</asp:ListItem>
                                </asp:DropDownList>
                                <asp:DropDownList ID="ddlAnoEncerrado" runat="server" />
                            </div>
                        </div>
                        <ajaxToolkit:TabContainer ID="TC03" runat="server" ActiveTabIndex="0" Width="100%">
                            <ajaxToolkit:TabPanel ID="TPCorrecaoEncerrados" runat="server" HeaderText="TabPanel4">
                                <HeaderTemplate>
                                    Pedidos Encerrados
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="menu_acoes">
                                        <div class="acoes">
                                            <ul>
                                                <li runat="server" style="width: 140px">
                                                    <asp:LinkButton class="iconRelatorio" ID="lnkVisContEnc" Text="Visual. Contabilização"
                                                        runat="server" />
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                    <div class="bordagrid">
                                        <asp:GridView ID="gridCorrecaoEncerrados" runat="server" CellPadding="4" ForeColor="#333333"
                                            GridLines="None" AutoGenerateColumns="False">
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
                                                        <input type="checkbox" id="chkTodos" onclick="selectAll(this, '#MainContent_TC01_TPEncerrado_TC03_TPCorrecaoEncerrados_gridCorrecaoEncerrados');" />
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:CheckBox ID="chkPedidoCorrecao" runat="server" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="Empresa" HeaderText="Empresa" />
                                                <asp:BoundField DataField="EndEmpresa" />
                                                <asp:BoundField DataField="Pedido" HeaderText="Pedido" />
                                                <asp:BoundField DataField="Classe" HeaderText="Classe" />
                                                <asp:BoundField DataField="Cliente" HeaderText="Cliente" />
                                                <asp:BoundField DataField="EndCliente" />
                                                <asp:BoundField DataField="Nome" />
                                                <asp:BoundField DataField="Passiva" HeaderText="Var. Passiva" />
                                                <asp:BoundField DataField="Ativa" HeaderText="Var.  Ativa" />
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:ImageButton ID="imgExtratoPedidoEnc" runat="server" Height="23px" ImageAlign="AbsMiddle"
                                                            ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imgExtratoPedidoEnc_Click"
                                                            data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido" Width="23px" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField />
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TPContabilizacaoEncerrados" runat="server" HeaderText="TabPanel4">
                                <HeaderTemplate>
                                    Estorno Contabilização
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="menu_acoes">
                                        <div class="acoes">
                                            <ul>
                                                <li runat="server" style="width: 140px">
                                                    <asp:LinkButton class="iconRelatorio" ID="lnkContEnc" Text="Contabilizar" runat="server" />
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                    <div class="bordagrid">
                                        <asp:GridView ID="gridContEncerrados" runat="server" CellPadding="4" ForeColor="#333333"
                                            GridLines="None" AutoGenerateColumns="False" Width="100%">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EditRowStyle BackColor="#999999" />
                                            <Columns>
                                                <asp:BoundField DataField="Pedido" HeaderText="Pedido" />
                                                <asp:BoundField DataField="Conta" HeaderText="Conta" />
                                                <asp:BoundField DataField="Cliente" HeaderText="Cliente" />
                                                <asp:BoundField DataField="EndCliente" />
                                                <asp:BoundField DataField="Nome" HeaderText="Nome" />
                                                <asp:BoundField DataField="DebitoOficial" HeaderText="Deb. Oficial" />
                                                <asp:BoundField DataField="creditooficial" HeaderText="Cred. Oficial" />
                                                <asp:BoundField DataField="Historico" HeaderText="Historico" />
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                        </ajaxToolkit:TabContainer>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
