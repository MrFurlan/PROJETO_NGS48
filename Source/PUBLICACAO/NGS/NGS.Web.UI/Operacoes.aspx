<%@ Page Title="" Language="vb" AutoEventWireup="True" MasterPageFile="~/Principal.Master"
    CodeBehind="Operacoes.aspx.vb" Inherits="NGS.Web.UI.Operacoes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 147px;
        }
    </style>
    <script type="text/javascript">
        $(document).ready(function () {
            $('#MainContent_TabContainer1_TabPanel2_ChkQuantidadePedido').click(function () {
                $('#MainContent_TabContainer1_TabPanel2_ChkQuantidadeFiscal').attr('checked', $(this).attr('checked') == 'checked');
            });

            $('#MainContent_TabContainer1_TabPanel2_ChkQuantidadeFiscal').click(function () {
                if ($('#MainContent_TabContainer1_TabPanel2_ChkQuantidadePedido').attr('checked') == 'checked') {
                    $('#MainContent_TabContainer1_TabPanel2_ChkQuantidadeFiscal').attr('checked', true);
                    msgbox('NÃO É POSSIVEL DESELECIONAR A "QUANTIDADE FISCAL" .FOI SELECIONADO "QUANTIDADE PEDIDO".', "ATENÇÃO!", "Info");
                }
            });
        });

    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngOperacao" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlConsultaPedidosXNotas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Cadastro de Operações
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0"
                Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Operações
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                    <div style="float: right;" >
                                        <li runat="server">
                                            <div class="row" style="margin-top: 0;">
                                                <div class="coltxt">
                                                    <asp:DropDownList ID="ddlUsuarios" runat="server" Width="175px" />
                                                </div>
                                            </div>
                                        </li>
                                    </div>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl" style="width: 113px;">
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtOperacaoCodigo" runat="server" Width="150px" MaxLength="2" data-ToolTip="default"
                                    ToolTip="Código de Cadastro da Operação." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl" style="width: 113px;">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtOperacaoDescricao" runat="server" Width="492px" MaxLength="50"
                                    data-ToolTip="default" ToolTip="Descrição da operação." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl" style="width: 113px;">
                                Classe:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlClasse" runat="server" Width="300px" DataMember="ClassesdeOperacoes"
                                    DataValueField="ClasseDescricao" DataTextField="ClasseDescricao">
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:CheckBox ID="ChkOperacaoProducao" runat="server" Text="Operação Especifica Para Estoque e Produção"
                                    data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:CheckBox ID="chkEstadoDestino" runat="server" Text="As configurações dos encargos dos documentos fiscais segue o Estado do Deposito de Destino e não do Cliente."
                                    data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="painelleft" style="width: 49.5%; margin-right: 4px;">
                            <div class="subtitulodiv">
                                Operações
                            </div>
                            <div class="bordagrid" style="margin-right: 4px;">
                                <asp:GridView ID="gridOperacao" runat="server" CellPadding="4" ForeColor="#333333"
                                    GridLines="None" OnSelectedIndexChanged="gridOperacao_SelectedIndexChanged" AutoGenerateColumns="False">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                        <asp:BoundField DataField="Codigo" HeaderText="Codigo" />
                                        <asp:BoundField DataField="Descricao" HeaderText="Descricao" />
                                        <asp:BoundField DataField="CodigoClasse" HeaderText="Classe" />
                                        <asp:CheckBoxField DataField="Producao" HeaderText="Producao" />
                                        <asp:CheckBoxField DataField="UFDepositoDestino" HeaderText="UF Dep.Destino">
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:CheckBoxField>
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
                        <div class="painelleft" style="width: 49.5%;">
                            <div class="subtitulodiv">
                                Sub-Operações
                            </div>
                            <div class="bordagrid">
                                <asp:GridView ID="gridSubOperacao" runat="server" CellPadding="4" ForeColor="#333333"
                                    GridLines="None" AutoGenerateColumns="False">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                        <asp:BoundField DataField="codigooperacao" HeaderText="Op." />
                                        <asp:BoundField DataField="codigo" HeaderText="Sub" />
                                        <asp:BoundField DataField="Classe" HeaderText="Classe" />
                                        <asp:BoundField DataField="Descricao" HeaderText="Descricao" />
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
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Sub-Operações
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnksubNovo" runat="server">
                                            <span>Gravar</span>
                                        </asp:LinkButton>
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnksubAtualizar" runat="server">
                                            <span>Atualizar</span>
                                        </asp:LinkButton>
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnksubExcluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;">
                                            <span>Excluir</span>
                                        </asp:LinkButton>
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnksubLimpar" runat="server">
                                            <span>Limpar</span>
                                        </asp:LinkButton>
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnksubRelatorio" runat="server">
                                            <span>Relatório</span>
                                        </asp:LinkButton>
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnksubAjuda" runat="server">
                                            <span>Ajuda</span>
                                        </asp:LinkButton>
                                    </li>
                                    <div style="float: right;">
                                        <li runat="server">
                                            <div class="row" style="margin-top: 0;">
                                                <div class="coltxt">
                                                    <asp:DropDownList ID="ddlSubUsuarios" runat="server" Width="175px" />
                                                </div>
                                            </div>
                                        </li>
                                    </div>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Op./Sub./Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigoOperacaoSuboperacao" runat="server" BackColor="#FFFFCC"
                                    placeholder="Op." onkeydown="return checkShortcut();" ReadOnly="True" Width="40px"
                                    data-ToolTip="default" ToolTip="Nº da operação, da sub-operação e descrição." />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigoSuboperacao" runat="server" MaxLength="2" placeholder="Sub."
                                    Width="40px" data-ToolTip="default" ToolTip="Nº da operação, da sub-operação e descrição." />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtDescricaoSubOperacao" runat="server" MaxLength="100" Width="485px"
                                    placeholder="Descrição" data-ToolTip="default" ToolTip="Nº da operação, da sub-operação e descrição." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Classe:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="CmbClasse" runat="server" DataMember="ClassesdeOperacoes" DataTextField="ClasseDescricao"
                                    DataValueField="ClasseDescricao" Width="250px" AutoPostBack="True">
                                </asp:DropDownList>
                            </div>
                            <div class="collbl">
                                Tipo:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RbtEntrada" runat="server" AutoPostBack="True" GroupName="entsai"
                                    Text="Entrada" data-ToolTip="default" ToolTip="Informar se é entrada ou saída." />
                                <asp:RadioButton ID="RbtSaida" runat="server" AutoPostBack="True" GroupName="entsai"
                                    Text="Saída" data-ToolTip="default" ToolTip="Informar se é entrada ou saída." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Situação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSituacao" runat="server" Width="250px" />
                            </div>
                            <div class="collbl">
                                Finalidade da Nota:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlFinalidadeDaNota" runat="server" Width="172px">
                                    <asp:ListItem Text=" " Value="0" />
                                    <asp:ListItem Text="1 - Normal" Value="1" />
                                    <asp:ListItem Text="2 - Complementar" Value="2" />
                                    <asp:ListItem Text="3 - Ajuste" Value="3" />
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            Comportamento da Sub-Operação
                        </div>
                        <asp:Panel ID="pnlchks" runat="server" Style="margin-top: 4px;">
                            <table style="width: 100%; border: 0px none;">
                                <tr>
                                    <td>
                                        <asp:CheckBox ID="chkPedido" runat="server" AutoPostBack="True" Text="Op. de Pedido"
                                            data-ToolTip="default" ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="ChkPrecoFixo" runat="server" Text="PreçoFixo" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="ChkDevolucao" runat="server" AutoPostBack="True" Text="Devolução"
                                            data-ToolTip="default" ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="ChkDeposito" runat="server" Text="Est.Arm.Terceiros" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="chkParaFinsDeExportacao" runat="server" Text="Para fins de Exportacao (Alfandegado)"
                                            data-ToolTip="default" ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:CheckBox ID="ChkUnitarioPedido" runat="server" Text="Unitário Pedido" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="chkLiminar" runat="server" Text="Liminar" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="ChkFinanceiro" runat="server" Text="Financeiro" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="chkProdutoDeTerceiros" runat="server" Text="Produto De Terceiros"
                                            data-ToolTip="default" ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="chkMemorando" runat="server" Text="Memorando Exportação" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:CheckBox ID="ChkQuantidadePedido" runat="server" Text="Quantidade Pedido" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="ChkContabil" runat="server" Text="Contábil" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="chkConsignacao" runat="server" Text="Consignacao" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="chkAmostraGratis" runat="server" Text="Amostra Grátis" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="chkNotaDebito" runat="server" Text="Nota de Débito" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td colspan="2">&nbsp;
                                    </td>
                                </tr>
                                <tr runat="server">
                                    <td colspan="5">&nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:CheckBox ID="ChkLaudo" runat="server" Text="Laudo" data-ToolTip="default" ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="ChkEstoqueInicial" runat="server" Text="Estoque Inicial" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="ChkQuantidadeFiscal" runat="server" Text=" Quantidade Fiscal" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="chkPecas" runat="server" Text="Peças / Meios" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td colspan="2">
                                        <asp:CheckBox ID="chkNotaCredito" runat="server" Text="Nota de Crédito" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:CheckBox ID="chkNumeroDoLote" runat="server" Text="Controlar Número do Lote" data-ToolTip="default" ToolTip="Controlar Número do Lote e Validade" />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="ChkEstoqueFisico" runat="server" Text="Estoque Físico" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="ChkQuantidadeFisica" runat="server" Text="Quantidade Física" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="chkCobraServico" runat="server" Text="Cobra Servico" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>&nbsp;
                                    </td>
                                </tr>
                                <tr>
                                    <td></td>
                                    <td colspan="2">
                                        <asp:CheckBox ID="ChkEstoqueFiscal" runat="server" Text="Estoque Fiscal" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td olspan="2">
                                        <asp:CheckBox ID="chkProprietarioDaMercadoria" runat="server" Text="Obriga Proprietario da Mercadoria" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                    <td>
                                        <asp:CheckBox ID="chkRepresentante" runat="server" Text="Obriga Representante" data-ToolTip="default"
                                            ToolTip="Cada tipo de operação será parametrizada marcando as opções desejadas." />
                                    </td>
                                </tr>
                            </table>
                        </asp:Panel>
                        <div class="row">
                            <div class="collbl">
                                Operação Contrapartida:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlOperacaoDestino" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlOperacaoDestino_SelectedIndexChanged"
                                    Width="300px">
                                </asp:DropDownList>
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSubOperacaoDestino" runat="server" Width="292px">
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Grupo de Contas:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlGrupoDeContas" runat="server" Width="596px">
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conta de Adiantamento:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlContaAdiantamento" runat="server" Width="596px">
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Apura Custos:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlTabelaDeApuracaoDeCustos" runat="server" Width="596px">
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Custos Contrapartida:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlContraPartidaCusto" runat="server" Width="596px">
                                </asp:DropDownList>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
