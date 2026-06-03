<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioEntradaSaidaNotas.aspx.vb" Inherits="NGS.Web.UI.RelatorioEntradaSaidaNotas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadRelatorioEntradaSaidaNotas() {
            $("#MainContent_lstEncargos").multiselect({
                header: "Escolha apenas 5 encargos!",
                selectedList: 5,
                click: function (e) {
                    if ($(this).multiselect("widget").find("input:checked").length > 5) {
                        return false;
                    }
                }
            }).multiselectfilter();

            $("#MainContent_LstPlanoDeCusto").multiselect({ selectedList: 5 }).multiselectfilter();
            $("#MainContent_lstClasseOp").multiselect().multiselectfilter();
            $("#MainContent_lstFinalidade").multiselect().multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadRelatorioEntradaSaidaNotas();
            var prmRelatorioEntradaSaidaNotas = Sys.WebForms.PageRequestManager.getInstance();
            prmRelatorioEntradaSaidaNotas.add_endRequest(pageLoadRelatorioEntradaSaidaNotas);
        });

        function downloadAndPrintPDF(apiUrl) {
            var printWindow = window.open(apiUrl);

            printWindow.onload = function () {
                printWindow.focus();

                // Atrasar a impressão para garantir que o PDF carregue completamente
                setTimeout(function () {
                    printWindow.print(); // Aciona a impressão
                }, 5000); // Ajuste o tempo se necessário
            };
        }

    </script>
    <style type="text/css">
        .collbl {
            width: 116px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngESNotas" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlESNotas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Movimentações Fiscais
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcelDados" runat="server" Text="Excel Dados" />
                                </li>
                            </ul>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbUnidadeNegocio" runat="server" AutoPostBack="True" Width="595px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" runat="server" Text="Empresa:"
                        data-ToolTip="default" ToolTip="Consolidar o cnpj da empresa." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEmpresa" runat="server" Width="595px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Isolar:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkPessoaFisica" runat="server" Text="Pessoa Fisica" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de pessoa para ser isolada." />
                    <asp:CheckBox ID="chkPessoaJuridica" runat="server" Text="Pessoa Juridica" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de pessoa para ser isolada." />
                    <asp:CheckBox ID="chkPessoaExterior" runat="server" Text="Pessoa Exterior" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de pessoa para ser isolada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente."
                        runat="server" Text="Cliente:" data-ToolTip="default" />
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtClientes" runat="server" Width="554px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaCliente" OnClick="cmdConsultaCliente_Click" runat="server"
                        Text=">" CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDeposito" runat="server" Width="554px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoDeposito" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaDeposito" OnClick="cmdConsultaDeposito_Click" runat="server"
                        CssClass="btn" Text=">" data-ToolTip="default" ToolTip="Local de armazenamento da mercadoria." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Origem/Dest:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="HOrigemDestino" runat="server" />
                    <asp:TextBox ID="txtOrigemDestino" Enabled="False" runat="server" Width="554px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnOrigemDestino" OnClick="btnOrigemDestino_Click" runat="server"
                        CssClass="btn" Text=">" data-ToolTip="default" ToolTip="Selecionar quanto ao cliente/fornecedor." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Representante:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliRepre" runat="server" />
                    <asp:TextBox ID="txtClienteRepresentante" runat="server" Enabled="False" Width="554px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaCliRepre" runat="server" CausesValidation="False" OnClick="cmdBuscaCliRepre_Click"
                        CssClass="btn" Text=">" UseSubmitBehavior="False" Width="24px" data-ToolTip="default"
                        ToolTip="Selecionar o representente quando houver." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Transportador:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoTransportador" runat="server" />
                    <asp:TextBox ID="txtClienteTransportador" runat="server" Enabled="False" Width="554px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaCliTransportador" runat="server" CausesValidation="False"
                        CssClass="btn" Text=">" UseSubmitBehavior="False" Width="24px" data-ToolTip="default"
                        ToolTip="Nome da pessoa/empresa responsável pelo transporte da mercadoria." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    PLano de Custo:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="LstPlanoDeCusto" runat="server" SelectionMode="Multiple" Width="593px"
                        data-ToolTip="default" ToolTip="Marcar conforme o tipo de custo." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Marca:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMarca" runat="server" Width="593px" AutoPostBack="True"
                        ToolTip="Selecionar a marca do produto." />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe Operacao:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstClasseOp" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="370px" data-ToolTip="default" ToolTip="Selecionar a classe de opeação desejada." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdComClasse" Text="Com Estas Classes" GroupName="cl" Checked="true"
                        runat="server" data-ToolTip="default" ToolTip="Selecionar a classe de opeação desejada." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdSemClasse" Text="Sem Estas Classes" GroupName="cl" runat="server"
                        data-ToolTip="default" ToolTip="Selecionar a classe de opeação desejada." />
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
                <div class="collbl">
                    Encargos:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstEncargos" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="593px" data-ToolTip="default" ToolTip="Selecionar os encargos envolvidos na operação." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Sit. Trib. ICMS:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlStICMS" runat="server" Width="592px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Sit. Trib. IPI:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlStIPI" runat="server" Width="592px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Sit. Trib. Pis/Cofins:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlStPISCOFINS" runat="server" Width="592px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Finalidade Pedido:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstFinalidade" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="370px" data-ToolTip="default" ToolTip="Objetivo da criação do pedido." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdComFinalidade" Text="Com Estas Finalidades" GroupName="fn"
                        Checked="true" runat="server" data-ToolTip="default" ToolTip="Objetivo da criação do pedido." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdSemFinalidade" Text="Sem Estas Finalidades" GroupName="fn"
                        runat="server" data-ToolTip="default" ToolTip="Objetivo da criação do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbSafra" runat="server" Width="592px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbSafra_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo CFOP:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbGrupoCFOP" runat="server" Width="592px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbGrupoCFOP_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSituacao" runat="server" Width="592px" />
                </div>
            </div>
            <div class="row" runat="server">
                <div class="coltxt" style="width: 100%;">
                    <asp:Panel ID="Panel3" runat="server" Width="100%" Height="100%">
                        <table style="width: 100%; height: 100%; border: 0px none; margin-bottom: 4px;">
                            <tr>
                                <td style="width: 48%;">
                                    <asp:Panel ID="pnlCfop" runat="server" Width="100%" Height="130px">
                                        <asp:ListBox ID="lstCfop" runat="server" Height="130px" Width="100%"></asp:ListBox>
                                    </asp:Panel>
                                </td>
                                <td align="center">
                                    <asp:ImageButton ID="imgAdicionar" OnClick="imgAdicionar_Click" runat="server" Width="15px"
                                        Height="15px" ImageUrl="~/images/ico-mais.gif" data-ToolTip="default" ToolTip="Adicionar CFOP para Lista"></asp:ImageButton><br />
                                    <br />
                                    <asp:ImageButton ID="imgRemover" OnClick="imgRemover_Click" runat="server" Width="15px"
                                        Height="15px" ImageUrl="~/images/ico-menos.gif" data-ToolTip="default" ToolTip="Remover CFOP da Lista"></asp:ImageButton>
                                </td>
                                <td style="width: 48%;">
                                    <asp:Panel ID="pnlSelecionados" runat="server" Width="100%" Height="130px">
                                        <asp:ListBox ID="lstCfopSelecionados" runat="server" Height="130px" Width="100%"></asp:ListBox>
                                    </asp:Panel>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkPeriodo" runat="server" AutoPostBack="True" Font-Bold="True"
                        Text="Data Movimento:" />
                </div>
                <div class="coltxt">
                    <asp:Panel ID="pnlDataMovimento" runat="server" Visible="False" HorizontalAlign="Left">
                        &nbsp;<asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="112px"
                            data-ToolTip="default" ToolTip="Selecionar o período desejado." />
                        &nbsp;a
                        <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="112px"
                            data-ToolTip="default" ToolTip="Selecionar o período desejado." />
                    </asp:Panel>
                </div>
            </div>
            <div class="row" runat="server">
                <div class="painelleft">
                    <div class="row">
                        <div class="collbl">
                            Tipo de Movimento:
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkEntrada" runat="server" AutoPostBack="True" Text="Entrada" Checked="True"
                                GroupName="ES" OnCheckedChanged="EntradaSaida_CheckedChanged" data-ToolTip="default"
                                ToolTip="Selecionar se é entrada ou saída." />
                            <asp:CheckBox ID="chkSaida" runat="server" AutoPostBack="True" Text="Saída" Checked="True"
                                GroupName="ES" OnCheckedChanged="EntradaSaida_CheckedChanged" data-ToolTip="default"
                                ToolTip="Selecionar se é entrada ou saída." />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Tipo de Relatório:
                        </div>
                        <div class="coltxt">
                            <asp:RadioButton ID="optFisico" runat="server" AutoPostBack="True" Text="Físico"
                                GroupName="TipoRelatorio" OnCheckedChanged="optFisico_CheckedChanged"
                                data-ToolTip="default" ToolTip="Selecionar se é relatório físico ou fiscal." />
                            <asp:RadioButton ID="optFiscal" runat="server" AutoPostBack="True" Text="Fiscal"
                                Checked="True" GroupName="TipoRelatorio" OnCheckedChanged="optFiscal_CheckedChanged" data-ToolTip="default"
                                ToolTip="Selecionar se é relatório físico ou fiscal." />
                            <asp:CheckBox ID="chkQuebraPorCliente" runat="server" Text="Quebra Por Cliente" data-ToolTip="default"
                                ToolTip="Selecionar se é relatório físico ou fiscal." />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Tipo de Frete:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlTipoFrete" runat="server">
                                <asp:ListItem>Todos</asp:ListItem>
                                <asp:ListItem>CIF</asp:ListItem>
                                <asp:ListItem>FOB</asp:ListItem>
                                <asp:ListItem Value="NEN">NENHUM</asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Pedido:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlTroca" runat="server">
                                <asp:ListItem>Todos</asp:ListItem>
                                <asp:ListItem>Só Troca</asp:ListItem>
                                <asp:ListItem>Sem Troca</asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Placa:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtPlaca" runat="server" Width="100px" MaxLength="8"
                                data-ToolTip="default" ToolTip="Inserir a placa do veículo." />
                        </div>
                    </div>
                </div>
                <div class="painelleft">
                    <div class="row">
                        <div class="collbl" style="width: 180px;">
                            Série:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtSerie" runat="server" Width="26px" data-ToolTip="default" ToolTip="Informar a série da nota." />
                        </div>
                    </div>
                    <div class="row">
                        <asp:Panel ID="pnlMedia" runat="server">
                            <div class="collbl" style="width: 180px;">
                                Média:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="optAritmetica" runat="server" GroupName="Media" Text="Aritmetica"
                                    data-ToolTip="default" ToolTip="Informar o tipo de média." />
                                <asp:RadioButton ID="optPonderada" runat="server" Checked="True" GroupName="Media"
                                    Text="Ponderada" data-ToolTip="default" ToolTip="Informar o tipo de média." />
                            </div>
                        </asp:Panel>
                        <asp:Panel ID="pnlOpcaoFiscal" runat="server" Visible="False">
                            <div class="collbl" style="width: 180px;">
                                OPÇÃO FISCAL(IGNORAR NOTAS):
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="chkGlobal" runat="server" Checked="false" Text="Global" />
                                <asp:CheckBox ID="chkRemessa" runat="server" Checked="false" Text="Remessa" />
                            </div>
                        </asp:Panel>
                    </div>
                    <div class="row">
                        <div class="collbl" style="width: 180px;">
                            <asp:CheckBox ID="chkCusto" runat="server" class="primario" Text="Somente Custo"
                                data-ToolTip="default" ToolTip="Apenas Nota Fiscal que Operação faz parte do Custo" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl" style="width: 180px;">
                            Troca De Nota:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlNfTroca" runat="server">
                                <asp:ListItem>Todos</asp:ListItem>
                                <asp:ListItem>Só Troca</asp:ListItem>
                                <asp:ListItem>Sem Troca</asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                </div>
                <div class="painelleft">
                    <div class="row">
                        <div class="collbl" style="width: 100px;">
                            Ordenada por:
                        </div>
                        <div class="coltxt">
                            <asp:RadioButton ID="RdNota" runat="server" Text="Nota" GroupName="Ordem" Font-Bold="False"
                                data-ToolTip="default" ToolTip="Selecionar se a ordem será por nota, nome ou movimento." />
                            <asp:RadioButton ID="RdNome" runat="server" Text="Nome" Checked="True" GroupName="Ordem"
                                Font-Bold="False" data-ToolTip="default" ToolTip="Selecionar se a ordem será por nota, nome ou movimento." />
                            <asp:RadioButton ID="RdMovimento" runat="server" Text="Movimento" GroupName="Ordem"
                                Font-Bold="False" data-ToolTip="default" ToolTip="Selecionar se a ordem será por nota, nome ou movimento." />
                        </div>
                    </div>
                    <div class="row">
                        <div id="divSeries" runat="server">
                            <div class="collbl" style="width: 100px;">
                                Listar séries D e F:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RadSerieSim" runat="server" Text="Sim" Checked="True" GroupName="series"
                                    data-ToolTip="default" ToolTip="Marcar se é para listar as séries D ou F." />
                                <asp:RadioButton ID="RadSerieNao" runat="server" Text="Não" GroupName="series" data-ToolTip="default"
                                    ToolTip="Marcar se é para listar as séries D ou F." />
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl" style="width: 100px;">
                            <asp:CheckBox ID="chkResumo" runat="server" AutoPostBack="True" Font-Bold="True"
                                Text="Resumo:" data-ToolTip="default" ToolTip="Marcar quando for somente resumo." />
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlTipoDeResumo" runat="server" Visible="False">
                                <asp:ListItem>Por Produto</asp:ListItem>
                                <asp:ListItem>Por CFOP</asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Num. do Registro:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" placeholder="Pedido" data-ToolTip="default"
                        ToolTip="Informar o número de pedido, CN pedido e contrato." />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCN" runat="server" placeholder="CN PEDIDO" data-ToolTip="default"
                        ToolTip="Informar o número de pedido, CN pedido e contrato." />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContrato" runat="server" placeholder="CONTRATO" data-ToolTip="default"
                        ToolTip="Informar o número de pedido, CN pedido e contrato." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Representante:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoRepresentante" runat="server" />
                    <asp:TextBox ID="txtRepresentante" runat="server" Enabled="False" Width="554px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnRepresentante" runat="server" OnClick="btnRepresentante_Click"
                        Text="&gt;" CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Nome do Representante." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Impressão NF:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkImpressao" runat="server" BorderStyle="None"
                        Text="Imprimir PDF" data-ToolTip="default" ToolTip="Imprimir notas fiscais" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkZIP" runat="server" BorderStyle="None"
                        Text="Zipar PDF" data-ToolTip="default" ToolTip="Zipar uma cópia dos PDFs das notas fiscais" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkAbrir" runat="server" BorderStyle="None"
                        Text="Abrir PDF" data-ToolTip="default" ToolTip="Abrir o PDF das notas fiscais" />
                </div>
                                <div class="collbl">
                    Página inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPaginaInicial" runat="server" CssClass="texto" MaxLength="3" Style="text-transform: uppercase"
                        TabIndex="8" Width="36px" data-ToolTip="default" ToolTip="Informar a página inicial" />
                </div>
                <div class="collbl">
                    Página final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPaginaFinal" runat="server" CssClass="texto" MaxLength="3" Style="text-transform: uppercase"
                        TabIndex="8" Width="36px" data-ToolTip="default" ToolTip="Informar a página final" />
                </div>
            </div>
             <div class="row">
                <div class="collbl colw">
                    Número/Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumeroNota" runat="server" TabIndex="7" CssClass="txtNumerico9"
                        Width="88px" data-ToolTip="default" ToolTip="Informar o número da nota." />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TextBox1" runat="server" CssClass="texto" MaxLength="3" Style="text-transform: uppercase"
                        TabIndex="8" Width="36px" data-ToolTip="default" ToolTip="Informar a série da nota." />
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 140px;">
                    <asp:CheckBox ID="chkAllTipos" data-ToolTip="default" ToolTip="Seleciona todos os Tipos de Documentos."
                        Text="Tipo De Documento:" runat="server" AutoPostBack="True" OnCheckedChanged="chkAllTipos_CheckedChanged" />
                </div>
                <div class="coltxt" style="line-height: 12px;">
                    <asp:CheckBoxList ID="chkTipoDeDocumento" runat="server" RepeatColumns="3" data-ToolTip="default"
                        ToolTip="Selecionar qual ou quais os tipos de documentos." />
                </div>
            </div>            
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
