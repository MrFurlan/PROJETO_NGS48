<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PedidosXItens.aspx.vb" Inherits="NGS.Web.UI.PedidosXItens" MaintainScrollPositionOnPostback="true" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1360px !important;
        }

        .txtNumerico {
            text-align: right;
        }
    </style>

    <script type="text/javascript">
        function pageLoadPed() {
            $("#MainContent_txtRegistro").keydown(function (event) {
                if (event.which == 13) {
                    event.preventDefault();
                    consulta();
                }
            });

            $("#txtDataPedido:enabled").datepicker({
                dateFormat: 'dd/mm/yy',
                dayNames: ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'],
                dayNamesMin: ['D', 'S', 'T', 'Q', 'Q', 'S', 'S', 'D'],
                dayNamesShort: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'],
                monthNames: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
                monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                nextText: 'Próximo',
                prevText: 'Anterior',
                showOn: "button",
                buttonImage: "Images/calendar.png",
                buttonImageOnly: true,
                showButtonPanel: true
            });

            $("#txtDataVencimentoPedido:enabled").datepicker({
                dateFormat: 'dd/mm/yy',
                dayNames: ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'],
                dayNamesMin: ['D', 'S', 'T', 'Q', 'Q', 'S', 'S', 'D'],
                dayNamesShort: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'],
                monthNames: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
                monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                nextText: 'Próximo',
                prevText: 'Anterior',
                showOn: "button",
                buttonImage: "Images/calendar.png",
                buttonImageOnly: true,
                showButtonPanel: true
            });

            $("#txtDataEntregaInicial:enabled").datepicker({
                dateFormat: 'dd/mm/yy',
                dayNames: ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'],
                dayNamesMin: ['D', 'S', 'T', 'Q', 'Q', 'S', 'S', 'D'],
                dayNamesShort: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'],
                monthNames: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
                monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                nextText: 'Próximo',
                prevText: 'Anterior',
                showOn: "button",
                buttonImage: "Images/calendar.png",
                buttonImageOnly: true,
                showButtonPanel: true
            });

            $("#txtDataEntregaFinal:enabled").datepicker({
                dateFormat: 'dd/mm/yy',
                dayNames: ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'],
                dayNamesMin: ['D', 'S', 'T', 'Q', 'Q', 'S', 'S', 'D'],
                dayNamesShort: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'],
                monthNames: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
                monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                nextText: 'Próximo',
                prevText: 'Anterior',
                showOn: "button",
                buttonImage: "Images/calendar.png",
                buttonImageOnly: true,
                showButtonPanel: true
            });

            $("#txtDataPedido").setMask("date");
            $("#txtDataVencimentoPedido").setMask("date");
            $("#txtDataEntregaInicial").setMask("date");
            $("#txtDataEntregaFinal").setMask("date");
        }

        $(document).ready(function () {
            pageLoadPed();
            var prmPedido = Sys.WebForms.PageRequestManager.getInstance();
            prmPedido.add_endRequest(pageLoadPed);
            setFocusRegistro();
        });

        function setFocusRegistro() {
            $("#MainContent_txtRegistro").focus();
        }

        function helpFixarDolar() {
            alert("Só fixe o dólar com autorização do comercial! Caso o dólar seja fixado, será desconsiderado o PTAX do dia!");
        }

        function downloadArquivo() {
            alert("Contrato Gerado.");
            $("#MainContent_imdDownload").click();
        };

        function consulta() {
            var btn = document.getElementById('<%=lnkConsultar.ClientID%>');
            btn.click();
        }

        function Arquivo() {
            $("#btnAdicionar").click();
        }

    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPedidosXItens" runat="server" AsyncPostBackTimeout="1000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPedidosXItens" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="txtCodigoCliente" runat="server" />
            <asp:HiddenField ID="txtCodigoPraca" runat="server" />
            <div class="titulodiv">
                Pedidos
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovoPedido" Text="Gravar" runat="server" OnClientClick=" if(!confirm('Deseja realmente incluir o Pedido?')) return false;" />
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" OnClientClick=" if(!confirm('Deseja realmente Alterar o Pedido?')) return false;" />
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkBloquear" Text="Bloquear" runat="server" OnClientClick="if(!confirm('Deseja realmente bloquear este Pedido?')) return false;" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconRelatorio rel" runat="server"><a>Reimprimir</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkContrato" Text="Contrato" runat="server" />
                        </li>
                        <li class="iconMail" runat="server">
                            <asp:LinkButton ID="lnkEnviarEmail" Text="Enviar E-mail" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                        <div style="float: right;">
                            <li runat="server">
                                <div class="row" style="margin-top: 0;">
                                    <div class="coltxt">
                                        <asp:Image ID="imgUsuario" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/man2.png"
                                            data-ToolTip="default" ToolTip="Usuário Lançamento" Width="20px" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlUsuarios" runat="server" Width="175px" />
                                    </div>
                                </div>
                            </li>
                            <li runat="server">
                                <div class="row" style="margin-top: 0;">
                                    <div class="coltxt">
                                        <asp:Button ID="btnFiscal" CssClass="btn" BackColor="Green" ForeColor="White" Height="26px"
                                            runat="server" OnClick="btnFiscal_Click" Text="Fiscal Aberto" UseSubmitBehavior="False" />
                                    </div>
                                </div>
                            </li>
                            <li runat="server">
                                <div class="row" style="margin-top: 0px;">
                                    <div class="coltxt">
                                        <asp:Button ID="btnFinanceiro" CssClass="btn" BackColor="Green" ForeColor="White"
                                            Height="26px" runat="server" OnClick="btnFinanceiro_Click" Text="Financ. Aberto"
                                            UseSubmitBehavior="False" />
                                    </div>
                                </div>
                            </li>
                            <li runat="server">
                                <div class="row" style="margin-top: 0;">
                                    <div class="coltxt">
                                        <asp:Button ID="btnBloqueado" runat="server" CssClass="btn" BackColor="Green" ForeColor="White"
                                            Height="26px" Text="Bloqueado" UseSubmitBehavior="False" Enabled="False" Visible="False" />
                                    </div>
                                </div>
                            </li>
                        </div>
                    </ul>
                </div>
            </div>
            <div class="painelleft" style="width: 19%;">
                <div class="row">
                    <div class="collbl">
                        Registro:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtRegistro" runat="server" CssClass="txtNumerico" Width="70px"
                            Enabled="False" data-ToolTip="default" ToolTip="Número de pedidos existente." />
                    </div>
                    <div class="coltxt">
                        <asp:ImageButton ID="imgExtratoPedido" runat="server" CssClass="imgconsultar" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                            OnClick="imgExtratoPedido_Click" data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Pedido (CN):
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPedido" runat="server" MaxLength="20" Width="99px" AutoPostBack="True"
                            data-ToolTip="default" ToolTip="Número da confirmação de negócio." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Contrato:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtContrato" runat="server" MaxLength="50" Width="99px" AutoPostBack="True"
                            data-ToolTip="default" ToolTip="Número do contrato vinculado ao pedido." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl w120" style="color: red;">
                        XPedNFe/nItemPed:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtXPedNFe" runat="server" MaxLength="15" Width="67px" AutoPostBack="True"
                            data-ToolTip="default" ToolTip="Esse campo se informado sai na TAG xPed na Nota Fiscal Eletrônica se for Nossa Emissão." />
                        <asp:TextBox ID="txtItemXPedNFe" runat="server" MaxLength="6" Width="20px" AutoPostBack="True"
                            data-ToolTip="default" ToolTip="Esse campo se informado sai na TAG nItemPed na Nota Fiscal Eletrônica se for Nossa Emissão. Só preencher se o Cliente quiser a númeração diferente do Item da Nota." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Data Pedido:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataPedido" ClientIDMode="Static" runat="server" AutoPostBack="True"
                            OnTextChanged="txtDataPedido_TextChanged" Width="75px" data-ToolTip="default" ToolTip="Data da criação do pedido." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Moeda:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="cmbMoeda" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbMoeda_SelectedIndexChanged"
                            Width="110px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Indexador:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="cmbIndexador" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbIndexador_SelectedIndexChanged"
                            Width="110px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Cotação Pedido:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtIndiceFixado" runat="server" AutoPostBack="True" CssClass="txtDecimal8"
                            OnTextChanged="txtIndiceFixado_TextChanged" Width="80px" data-ToolTip="default"
                            ToolTip="Valor da cotação diária do dólar." />
                    </div>
                    <div class="coltxt">
                        <img src="Images/question.jpg" alt="" width="22" height="22" id="imgAjudaFixarDolar1"
                            onclick="msgbox('Só fixe o dólar com autorização do comercial! Caso o dólar seja fixado, será desconsiderado o PTAX do dia!', 'ATENÇÃO!', 'Info');"
                            runat="server" style="cursor: pointer" />
                    </div>
                </div>
            </div>
            <div class="painelleft" style="width: 28%;">
                <div class="row">
                    <div class="collbl">
                        Situação:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="cmbSituacao" runat="server" Width="240px" Enabled="False" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Finalidade:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="cmbFinalidade" runat="server" Width="240px" AutoPostBack="True" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Safra:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="cmbSafra" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbSafra_SelectedIndexChanged"
                            Width="240px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Taxa de Juro:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtTaxa" runat="server" AutoPostBack="True" CssClass="txtDecimal"
                            OnTextChanged="txtTaxa_TextChanged" Width="50px" data-ToolTip="default" ToolTip="Juro a ser cobrado após o vencimento da safra." />
                    </div>
                    <div class="coltxt">
                        <asp:Label ID="Label5" runat="server" Font-Bold="True" Text="a.a." />
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkAgruparFinanceiro" runat="server" AutoPostBack="True" Font-Bold="True"
                            OnCheckedChanged="chkAgruparFinanceiro_CheckedChanged" Text="Agrupar Financeiro"
                            Visible="False" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Validade:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataVencimentoPedido" runat="server" AutoPostBack="True" ClientIDMode="Static" OnTextChanged="txtDataVencimentoPedido_TextChanged"
                            Width="75px" data-ToolTip="default" ToolTip="Refere-se ao prazo de final da safra." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Periodo Entrega:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataEntregaInicial" runat="server" AutoPostBack="True" ClientIDMode="Static" OnTextChanged="txtDataEntregaInicial_TextChanged"
                            Width="75px" data-ToolTip="default" ToolTip="Refere-se à data de registro do pedido até o fim da safra." />
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataEntregaFinal" runat="server" AutoPostBack="True" ClientIDMode="Static" OnTextChanged="txtDataEntregaFinal_TextChanged"
                            Width="75px" data-ToolTip="default" ToolTip="Refere-se à data de registro do pedido até o fim da safra." />
                    </div>
                </div>
                <div class="row">
                    <div class="coltxt">
                        <asp:CheckBox ID="chkIndexadorFixo" runat="server" AutoPostBack="True" Text="Fixar Cotação p/ NF" Enabled="true" />
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlComercializacao" runat="server" AutoPostBack="True" Width="114px"
                            Style="margin-left: 7px;">
                            <asp:ListItem Value="0">Industria</asp:ListItem>
                            <asp:ListItem Selected="True" Value="1">Comércio</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                    <div class="coltxt" style="line-height: 16px;">
                        <asp:CheckBox ID="chkPeso" runat="server" AutoPostBack="True" Text="Vale o Peso do Cliente / Fornecedor" data-ToolTip="default" ToolTip="Selecionar se o peso informado pelo cliente/fornecedor será o considerado." />
                        <br />
                        <asp:CheckBox ID="chkTemVariacao" runat="server" AutoPostBack="True" Text="Sofre Variação Cambial"
                            Visible="False" />
                    </div>
                </div>
                <div class="row" id="divNaviosXInvoice" runat="server" visible="False">
                    <div class="collbl">
                        Invoice:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtNaviosXInvoice" runat="server" Width="209px" Enabled="False" ToolTip="Código Invoice e descrição do navio." />
                        <asp:LinkButton class="iconConsultar" ID="lnkConsultarNaviosXInvoice" runat="server" />
                    </div>
                </div>
            </div>
            <div class="painelleft" style="width: 53%;">
                <div class="row">
                    <div class="collbl">
                        Unidade:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="cmbUnidadeNegocio" runat="server" Enabled="false" AutoPostBack="True" OnSelectedIndexChanged="cmbUnidadeNegocio_SelectedIndexChanged"
                            Width="584px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Empresa:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="cmbEmpresa" runat="server" Enabled="false" AutoPostBack="True" OnSelectedIndexChanged="cmbEmpresa_SelectedIndexChanged"
                            Width="584px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Op. Com:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="cmbOperacao" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbOperacao_SelectedIndexChanged"
                            Width="287px" />
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="cmbSubOperacao" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbSubOperacao_SelectedIndexChanged"
                            Width="293px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Cliente / Fornec.:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtCliente" runat="server" Width="543px" Enabled="False" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="cmdConsultarCliente" OnClick="cmdConsultarCliente_Click" runat="server"
                            CssClass="btn" UseSubmitBehavior="False" Text=">" data-ToolTip="default" ToolTip="Nome da empresa destinatária/remetente do pedido." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Pr Pagto:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPraca" runat="server" Enabled="False" Width="543px" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnConsultaPraca" runat="server" OnClick="btnConsultaPraca_Click"
                            CssClass="btn" Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Praça de Pagamento - Informações para depósito bancário." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Dados Bancários:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDadosBancarios" runat="server" Enabled="False" Width="543px" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="cmdConsultaDadosBancarios" runat="server" OnClick="cmdConsultaDadosBancarios_Click"
                            CssClass="btn" Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Relação das contas do cliente/fornecedor bem como os dados bancários." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        <asp:Label ID="lblCifFob" runat="server" Text="Frete" />
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rbCIF" runat="server" AutoPostBack="True" Checked="True" GroupName="CIF/FOB/TERCEIRO/NENHUM"
                            Text="(CIF)" data-ToolTip="default" ToolTip="Selecionar se o frete será por conta da empresa, cliente/fornecedor, terceiros ou nenhum." />
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rbFOB" runat="server" AutoPostBack="True" GroupName="CIF/FOB/TERCEIRO/NENHUM"
                            Text="(FOB)" data-ToolTip="default" ToolTip="Selecionar se o frete será por conta da empresa, cliente/fornecedor, terceiros ou nenhum." />
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rbTER" runat="server" AutoPostBack="True" GroupName="CIF/FOB/TERCEIRO/NENHUM"
                            Text="(TERCEIRO)" data-ToolTip="default" ToolTip="Selecionar se o frete será por conta da empresa, cliente/fornecedor, terceiros ou nenhum." />
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rbNenhum" runat="server" AutoPostBack="True" GroupName="CIF/FOB/TERCEIRO/NENHUM"
                            Text="(NENHUM)" data-ToolTip="default" ToolTip="Selecionar se o frete será por conta da empresa, cliente/fornecedor, terceiros ou nenhum." />
                    </div>
                    <div class="collbl">
                        <asp:Label ID="lblTipoOperacao" runat="server" Text="Tipo Operação" />
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkTroca" runat="server" AutoPostBack="True" Text="Troca" data-ToolTip="default" ToolTip="Selecionar uma das opções (troca, antecipada, recompra) de acordo com a particularidade do pedido." />
                        <asp:CheckBox ID="chkAntecipada" runat="server" AutoPostBack="True" Text="Antecipada" data-ToolTip="default" ToolTip="Selecionar uma das opções (troca, antecipada, recompra) de acordo com a particularidade do pedido." />
                        <asp:CheckBox ID="chkRecompra" runat="server" AutoPostBack="True" Text="Recompra" data-ToolTip="default" ToolTip="Selecionar uma das opções (troca, antecipada, recompra) de acordo com a particularidade do pedido." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Embalagem:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="cmbEmbalagem" runat="server"
                            Width="150px" AutoPostBack="True" />
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rbPagAnterior" runat="server" Checked="False" GroupName="PAGAMENTO" AutoPostBack="True"
                            Text="Ant. Pag." data-ToolTip="default" ToolTip="Selecionar se o carregamento é anterior ao pagamento" />
                        <asp:RadioButton ID="rbPagPosterior" runat="server" Checked="False" GroupName="PAGAMENTO" AutoPostBack="True"
                            Text="Post. Pag." data-ToolTip="default" ToolTip="Selecionar se o carregamento é posterior ao pagamento" />
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rbPtaxAnterior" runat="server" Checked="False" GroupName="PTAX" AutoPostBack="True"
                            Text="Ant. Ptax Pag." data-ToolTip="default" ToolTip="Selecionar se o Ptax é anterior ao pagamento" />
                        <asp:RadioButton ID="rbPtaxPosterior" runat="server" Checked="False" GroupName="PTAX" AutoPostBack="True"
                            Text="Ant. Ptax Car." data-ToolTip="default" ToolTip="Selecionar se o Ptax é anterior ao carregamento" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Local de Embarque:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtLocalDeEmbarque" runat="server" MaxLength="40" Width="200px"
                            data-ToolTip="default" ToolTip="Local de embarque." />
                    </div>
                </div>
            </div>
            <div class="row" id="divContratoArquivo" runat="server" visible="false">
                <div class="collbl">
                    Arquivo:
                </div>
                <div class="coltxt">
                    <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                        <ContentTemplate>
                            <asp:TextBox ID="txtArquivoDeSaida" runat="server" Width="500px" Enabled="False" />
                            <asp:ImageButton ID="imdDownload" runat="server" ImageAlign="AbsMiddle" ImageUrl="Images/download32x32.png"
                                Style="margin-top: 0;" OnClick="imdDownload_Click" data-ToolTip="default" ToolTip="Baixar Arquivo"
                                Height="22px" Width="22px" />
                        </ContentTemplate>
                        <Triggers>
                            <asp:PostBackTrigger ControlID="imdDownload" />
                        </Triggers>
                    </asp:UpdatePanel>
                </div>
            </div>
            <ajaxToolkit:TabContainer ID="tcPedido" runat="server" ActiveTabIndex="0" Width="100%"
                Style="clear: both;">
                <ajaxToolkit:TabPanel runat="server" HeaderText="Produto" ID="TabProduto">
                    <HeaderTemplate>
                        Produto
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes" runat="server">
                            <div class="acoes">
                                <ul>
                                    <li class="iconMais" runat="server">
                                        <asp:LinkButton ID="lnkAdicionarItem" Text="Adicionar Item" runat="server" data-ToolTip="default" ToolTip="Inserir produtos ao pedido." />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 82%;">
                            <div class="bordagrid" style="height: 225px;">
                                <asp:GridView ID="gridItensPedido" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:TemplateField HeaderText="Alterar">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lnkComplementoEstorno" runat="server" OnClick="lnkComplementoEstorno_Click">
                                                        <img src="Images/detalhes.png" alt="Complemento / Estorno do Item" />
                                                </asp:LinkButton>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" />
                                        <asp:BoundField DataField="Descricao" HeaderText="Nome" HeaderStyle-HorizontalAlign="Left" />

                                        <asp:TemplateField HeaderText="Qtde Pedido">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                            <ItemTemplate>
                                                <div style="float: right;">
                                                    <div style="float: left;">
                                                        <asp:Literal ID="lblQtdeFat" runat="server" Text='<%# String.Format("{0:N4}",eval("QuantidadePedidoFaturamento")) %>' />
                                                        <br />
                                                        <asp:Literal ID="lblQtdeCom" runat="server" Text='<%# String.Format("{0:N4}",eval("QuantidadePedidoComercializacao")) %>' />
                                                    </div>
                                                    <div style="float: right; padding-left: 6px;">
                                                        <asp:Literal ID="lblUnFat" runat="server" Text='<%# Eval("UnidadeFaturamento")%>' />
                                                        <br />
                                                        <asp:Literal ID="lblUnCom" runat="server" Text='<%# Eval("CodigoUnidadeComercializacao")%>' />
                                                    </div>
                                                </div>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Unitario Medio">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                            <ItemTemplate>
                                                <asp:Literal ID="lblUnMdFat" runat="server" Text='<%# String.Format("{0:N10}",eval("UnitarioMedioFaturamento")) %>' />
                                                <br />
                                                <asp:Literal ID="lblUnMdCom" runat="server" Text='<%# String.Format("{0:N10}",eval("UnitarioMedioComercializacao")) %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Valor">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                            <ItemTemplate>
                                                <div style="float: right;" runat="server">
                                                    <div style="float: right;" runat="server">
                                                        <asp:Literal ID="PVOC" runat="server" Text='<%# eval("SaldoItem.CifraoOficial") %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="PVMC" runat="server" Text='<%# eval("SaldoItem.CifraoPedido") %>'></asp:Literal>
                                                    </div>
                                                    <div style="float: right; padding-right: 6px;">
                                                        <asp:Literal ID="PVO" runat="server" Text='<%# String.Format("{0:N2}",eval("Lancamentos.TotalOficialPrd")) %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="PVM" runat="server" Text='<%# String.Format("{0:N2}",eval("Lancamentos.TotalMoedaPrd"))%>'></asp:Literal>
                                                    </div>

                                                </div>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Qtde Entregue">
                                            <ItemTemplate>
                                                <div style="float: right;" runat="server">

                                                    <div style="float: Left; padding-right: 6px;">
                                                        <asp:Literal ID="Literal1" runat="server" Text="Unid/Moeda"></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="LD" runat="server" Text="Direta"></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="LG" runat="server" Text="Futura"></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="LR" runat="server" Text="Remessa"></asp:Literal>
                                                    </div>
                                                    <div style="float: Left;">
                                                        <asp:Literal ID="QEF" runat="server" Text='<%# eval("UnidadeFaturamento") %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L17F" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.QtdeEntregueFiscalDireta")) %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L18F" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.QtdeEntregueFiscalGlobal"))%>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L19F" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.QtdeEntregueFiscalRemessa")) %>'></asp:Literal>
                                                    </div>
                                                    <div style="float: right;" runat="server">
                                                        <asp:Literal ID="QEC" runat="server" Text='<%# Eval("CodigoUnidadeComercializacao")%>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L17C" runat="server" Text='<%# String.Format("{0:N2}", Eval("SaldoItem.QtdeEntregueFiscalDireta") / Eval("UnidadeComercializacaoFatorDeConversao"))%>' />
                                                        <br />
                                                        <asp:Literal ID="L18C" runat="server" Text='<%# String.Format("{0:N2}", Eval("SaldoItem.QtdeEntregueFiscalGlobal") / Eval("UnidadeComercializacaoFatorDeConversao"))%>' />
                                                        <br />
                                                        <asp:Literal ID="L19C" runat="server" Text='<%# String.Format("{0:N2}", Eval("SaldoItem.QtdeEntregueFiscalRemessa") / Eval("UnidadeComercializacaoFatorDeConversao"))%>' />
                                                    </div>
                                                </div>
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Vlr Entregue">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                            <ItemTemplate>
                                                <div runat="server" style="float: right;">
                                                    <div runat="server" style="float: right;">
                                                        <asp:Literal ID="L08C" runat="server" Text='<%# eval("SaldoItem.CifraoPedido") %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L08" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaMoedaDiretaBruto")) %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L09" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaMoedaGlobalBruto")) %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L10" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaMoedaRemessaBruto")) %>'></asp:Literal>
                                                    </div>
                                                    <div style="float: right; padding-right: 4px;">
                                                        <asp:Literal ID="L05C" runat="server" Text='<%# eval("SaldoItem.CifraoOficial") %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L05" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaOficialDiretaBruto")) %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L06" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaOficialGlobalBruto")) %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L07" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaOficialRemessaBruto")) %>'></asp:Literal>
                                                    </div>
                                                </div>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Qtde Saldo">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                            <ItemTemplate>
                                                <div style="float: right;" runat="server">

                                                    <div style="float: right;" runat="server">
                                                        <asp:Literal ID="QSC" runat="server" Text='<%# eval("CodigoUnidadeComercializacao") %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L20C" runat="server" Text='<%# String.Format("{0:N4}", Eval("SaldoItem.SaldoQtdeDiretoFiscal") / Eval("UnidadeComercializacaoFatorDeConversao"))%>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L21C" runat="server" Text='<%# String.Format("{0:N4}", Eval("SaldoItem.SaldoQtdeGlobalFiscal") / Eval("UnidadeComercializacaoFatorDeConversao"))%>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L22C" runat="server" Text='<%#String.Format("{0:N4}", Eval("SaldoItem.SaldoQtdeRemessaFiscal") / Eval("UnidadeComercializacaoFatorDeConversao"))%>'></asp:Literal>
                                                    </div>
                                                    <div style="float: left;">
                                                        <asp:Literal ID="QSF" runat="server" Text='<%# eval("UnidadeFaturamento") %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L20F" runat="server" Text='<%# String.Format("{0:N4}", Eval("SaldoItem.SaldoQtdeDiretoFiscal")) %>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L21F" runat="server" Text='<%# String.Format("{0:N4}", Eval("SaldoItem.SaldoQtdeGlobalFiscal"))%>'></asp:Literal>
                                                        <br />
                                                        <asp:Literal ID="L22F" runat="server" Text='<%# String.Format("{0:N4}", Eval("SaldoItem.SaldoQtdeRemessaFiscal")) %>'></asp:Literal>
                                                    </div>
                                                </div>

                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Vlr Saldo">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                            <ItemTemplate>
                                                <div runat="server" style="float: right;">
                                                    <asp:Literal ID="L14C" runat="server" Text='<%# eval("SaldoItem.CifraoPedido") %>'></asp:Literal>
                                                    <br />
                                                    <asp:Literal ID="L14" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorMoedaGlobalDireto")) %>'></asp:Literal>
                                                    <br />
                                                    <asp:Literal ID="L15" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorMoedaGlobalDireto")) %>'></asp:Literal>
                                                    <br />
                                                    <asp:Literal ID="L16" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorMoedaRemessa")) %>'></asp:Literal>
                                                </div>
                                                <div style="float: right; padding-right: 4px;">
                                                    <asp:Literal ID="L11C" runat="server" Text='<%# eval("SaldoItem.CifraoOficial") %>'></asp:Literal>
                                                    <br />
                                                    <asp:Literal ID="L11" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorOficialGlobalDireto")) %>'></asp:Literal>
                                                    <br />
                                                    <asp:Literal ID="L12" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorOficialGlobalDireto")) %>'></asp:Literal>
                                                    <br />
                                                    <asp:Literal ID="L13" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorOficialRemessa")) %>'></asp:Literal>
                                                </div>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Encargos">
                                            <ItemTemplate>
                                                <div class="row">
                                                    <asp:LinkButton ID="lnkEncargos" runat="server" OnClick="lnkEncargos_Click"><img 
                                                                            src="Images/ico-mais.gif" alt="Encargos do Produto" /> </asp:LinkButton>
                                                </div>
                                                <div class="row">
                                                    <asp:Label ID="LRT" runat="server" Text="Com Retenção" Font-Bold="True" ForeColor="Blue"></asp:Label>
                                                </div>
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
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
                                <asp:GridView ID="gridItensPedidoDeposito" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:TemplateField HeaderText="Comp./Estorno">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lnkComplementoEstorno" runat="server" OnClick="lnkComplementoEstorno_Click"><img src="Images/detalhes.png" 
                                                                            alt="Complemento/Estorno do Item" /></asp:LinkButton>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" />
                                        <asp:BoundField DataField="Descricao" HeaderText="Nome" />
                                        <asp:BoundField DataField="CodigoUnidadeComercializacao" HeaderText="Und">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="QuantidadePedidoComercializacao" DataFormatString="{0:N4}"
                                            HeaderText="Qtde Pedido">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="UnitarioMedio" DataFormatString="{0:N10}" HeaderText="Unit.Medio">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PedidoValor" DataFormatString="{0:N2}" HeaderText="Valor">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="SaldoItem.QtdeContratadoFiscal" DataFormatString="{0:N2}"
                                            HeaderText="Qtde Contratada" />
                                        <asp:BoundField DataField="SaldoItem.QtdeEntregueFiscalDeposito" DataFormatString="{0:N2}"
                                            HeaderText="Qtde Entregue" />
                                        <asp:BoundField DataField="SaldoItem.VlrNotaOficialDepositoBruto" DataFormatString="{0:N2}"
                                            HeaderText="Vlr Entregue Oficial" />
                                        <asp:BoundField DataField="SaldoItem.VlrNotaMoedaDepositoBruto" DataFormatString="{0:N2}"
                                            HeaderText="Vlr Entregue Moeda" />
                                        <asp:TemplateField HeaderText="Encargos">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lnkEncargos" runat="server" OnClick="lnkEncargos_Click"><img 
                                                                            src="Images/ico-mais.gif" alt="Encargos do Produto" /> </asp:LinkButton>
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
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
                                <asp:GridView ID="gridItensPedidoAFixar" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:TemplateField HeaderText="Comp./Estorno">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lnkComplementoEstorno" runat="server" OnClick="lnkComplementoEstorno_Click"><img src="Images/detalhes.png" 
                                                                            alt="Complemento/Estorno do Item" /></asp:LinkButton>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" />
                                        <asp:BoundField DataField="Descricao" HeaderText="Nome" />
                                        <asp:BoundField DataField="CodigoUnidadeComercializacao" HeaderText="Und">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="QuantidadePedidoComercializacao" DataFormatString="{0:N4}"
                                            HeaderText="Qtde Pedido">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="UnitarioMedio" DataFormatString="{0:N10}" HeaderText="Unit.Medio">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="PedidoValor" DataFormatString="{0:N2}" HeaderText="Valor">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="SaldoItem.QtdeContratadoFiscal" DataFormatString="{0:N2}"
                                            HeaderText="Qtde Contratado" />
                                        <asp:BoundField DataField="SaldoItem.QtdeEntregueFiscalAFixar" DataFormatString="{0:N2}"
                                            HeaderText="Qtde Entregue" />
                                        <asp:BoundField DataField="SaldoItem.VlrNotaOficialAFixarBruto" DataFormatString="{0:N2}"
                                            HeaderText="Vlr Entregue R$" />
                                        <asp:BoundField DataField="SaldoItem.VlrNotaMoedaAFixarBruto" DataFormatString="{0:N2}"
                                            HeaderText="Vlr Entregue $" />
                                        <asp:BoundField DataField="SaldoItem.QtdeFixacao" DataFormatString="{0:N2}" HeaderText="Qtde Fixada" />
                                        <asp:BoundField DataField="SaldoItem.VlrFixacaoOficial" DataFormatString="{0:N2}"
                                            HeaderText="Vlr Fixado R$" />
                                        <asp:BoundField DataField="SaldoItem.VlrFixacaoMoeda" DataFormatString="{0:N2}" HeaderText="Vlr Fixado $" />
                                        <asp:BoundField DataField="SaldoItem.SaldoQtdeAFixar" DataFormatString="{0:N2}" HeaderText="Saldo Qtde AFixar" />
                                        <asp:TemplateField HeaderText="Encargos">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lnkEncargos" runat="server" OnClick="lnkEncargos_Click"><img 
                                                                            src="Images/ico-mais.gif" alt="Encargos do Produto" /> </asp:LinkButton>
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Fixações">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lnkFixacoes" runat="server" OnClick="lnkFixacoes_Click"><img 
                                                                            src="Images/ico-mais.gif" alt="Fixações do Produto" /></asp:LinkButton>
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
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
                <ajaxToolkit:TabPanel runat="server" HeaderText="Depósito" ID="TabDeposito">
                    <HeaderTemplate>
                        Depósito
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconMais" runat="server">
                                        <asp:LinkButton ID="lnkAdicionarDeposito" Text="Adicionar Depósito" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tipo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTipoDeposito" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlTipoDeposito_SelectedIndexChanged"
                                    ToolTip="Tipo do depósito."
                                    Width="230px" />
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="chkDepPrincipal" runat="server" Text="Principal" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Deposito:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbDepositos" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbDepositos_SelectedIndexChanged"
                                    Width="745px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnBuscarDeposito" runat="server" OnClick="btnBuscarDeposito_Click"
                                    CssClass="btn" Text=" &gt; " UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Nome do depósito referente ao cliente/fornecedor ou empresa." />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 150px;">
                            <asp:GridView ID="GridDepositos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="GridDepositos_SelectedIndexChanged"
                                Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True">
                                        <HeaderStyle Width="10px" />
                                        <ItemStyle Width="10px" CssClass="tooltip" />
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Tipo" HeaderText="Tipo">
                                        <HeaderStyle HorizontalAlign="Left" Width="60px" />
                                        <ItemStyle HorizontalAlign="Left" Width="60px" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Principal">
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkDepositoPrincipal" runat="server" Checked='<%# Eval("Principal") %>'
                                                Enabled="False" />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Center" Width="60px" />
                                        <ItemStyle HorizontalAlign="Center" Width="60px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Codigo">
                                        <HeaderStyle HorizontalAlign="Left" Width="110px" />
                                        <ItemStyle HorizontalAlign="Left" Width="110px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoEndereco" HeaderText="End.">
                                        <HeaderStyle HorizontalAlign="Left" Width="10px" />
                                        <ItemStyle HorizontalAlign="Left" Width="10px" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Nome">
                                        <ItemTemplate>
                                            <asp:Label ID="Label1" runat="server" Text='<%# Bind("Deposito.Nome") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Endereço">
                                        <ItemTemplate>
                                            <asp:Label ID="Label2" runat="server" Text='<%# Bind("Deposito.Endereco") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Cidade">
                                        <ItemTemplate>
                                            <asp:Label ID="Label3" runat="server" Text='<%# Bind("Deposito.Cidade") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Estado">
                                        <ItemTemplate>
                                            <asp:Label ID="Label4" runat="server" Text='<%# Bind("Deposito.CodigoEstado") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirDeposito" runat="server" Height="18px" ImageUrl="~/Images/erro.jpg"
                                                OnClick="imgExcluirDeposito_Click" Style="cursor: pointer" data-ToolTip="default"
                                                ToolTip="Excluir Depósito" Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                        <asp:Panel ID="pnlRoteiro" runat="server" Width="100%" Height="100%" Visible="False">
                            <div class="subtitulodiv">
                                Roteiros
                            </div>
                            <div class="menu_acoes">
                                <div class="acoes">
                                    <ul>
                                        <li class="iconMais" runat="server">
                                            <asp:LinkButton ID="lnkGerarRoteiro" Text="Gerar" runat="server" />
                                        </li>
                                        <li class="iconLimpar" runat="server">
                                            <asp:LinkButton ID="lnkLimparRoteiro" Text="Limpar" runat="server" />
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div class="painelleft" style="width: 79%;">
                                <div class="bordagrid" style="height: 200px;">
                                    <asp:GridView ID="grdRoteiros" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                        ForeColor="#333333" GridLines="None" Height="100%" Width="100%">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:TemplateField HeaderText="Tipo" SortExpression="TipoOrigem">
                                                <ItemTemplate>
                                                    <asp:Label ID="lblTipoOrigem" runat="server" Text='<%# Eval("TipoOrigem")%>' />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Nome">
                                                <ItemTemplate>
                                                    <asp:Label ID="lblCodigoOrigem" runat="server" Text='<%# Eval("Origem.Nome")%>' />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Endereço">
                                                <ItemTemplate>
                                                    <asp:Label ID="lblEnderecoOrigem" runat="server" Text='<%# Eval("Origem.Endereco")%>' />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Tipo">
                                                <ItemTemplate>
                                                    <asp:Label ID="lblTipoDestino" runat="server" Text='<%# Eval("TipoDestino")%>' />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Nome">
                                                <ItemTemplate>
                                                    <asp:Label ID="lblCodigoDestino" runat="server" Text='<%# Eval("Destino.Nome")%>' />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Endereço">
                                                <ItemTemplate>
                                                    <asp:Label ID="lblEnderecoDestino" runat="server" Text='<%# Eval("Destino.Endereco")%>' />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Via de Transporte">
                                                <ItemTemplate>
                                                    <asp:DropDownList ID="ddlTabelaDeViaDeTransporte" runat="server" AutoPostBack="true"
                                                        DataSource='<%# Bind("TabelaViaDeTransporte")%>' DataTextField="Descricao" DataValueField="Codigo"
                                                        OnSelectedIndexChanged="ddlTabelaDeViaDeTransporte_SelectedIndexChanged" SelectedValue='<%# Bind("CodigoViaDeTransporte")%>'
                                                        Width="95px" />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Original">
                                                <ItemTemplate>
                                                    <asp:TextBox ID="txtQuantidadeOriginal" runat="server" CssClass="txtDecimal" Enabled="false"
                                                        Style="text-align: right; width: 60px" Text='<%# Bind("QuantidadeOriginal", "{0:n2}")%>' />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Atual">
                                                <ItemTemplate>
                                                    <asp:TextBox ID="txtQuantidadeAtual" runat="server" AutoPostBack="true" CssClass="txtDecimal"
                                                        OnTextChanged="txtQuantidadeAtual_TextChanged" Style="text-align: right; width: 60px"
                                                        Text='<%# Bind("QuantidadeAtual", "{0:n2}")%>' />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Por Kg">
                                                <ItemTemplate>
                                                    <asp:TextBox ID="txtValor" runat="server" AutoPostBack="true" CssClass="txtDecimal9"
                                                        OnTextChanged="txtValor_TextChanged" Style="text-align: right; width: 60px"
                                                        Text='<%# Bind("Valor", "{0:n9}")%>' />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Por Ton">
                                                <ItemTemplate>
                                                    <asp:TextBox ID="txtValorPorTonelada" runat="server" AutoPostBack="true" CssClass="txtDecimal2"
                                                        OnTextChanged="txtValor_TextChanged" Style="text-align: right; width: 60px"
                                                        Text='<%# Bind("ValorPorTonelada", "{0:n2}")%>' Enabled="false" />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Log">
                                                <ItemTemplate>
                                                    <asp:ImageButton ID="imgRoteiroLog" runat="server" CssClass="imgconsultar" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                        data-ToolTip="default" ToolTip="Visualizar Log de Alterações de Valor" OnClick="imgRoteiroLog_Click" />
                                                </ItemTemplate>
                                            </asp:TemplateField>
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
                            <div class="painelright" id="divTotalRoteiro" style="width: 20%; height: 100%;" runat="server"
                                visible="False">
                                <div class="subtitulodiv">
                                    <label>
                                        Resumo do Roteiro
                                    </label>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Embarque:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtTotalQtdEmbarque" runat="server" CssClass="txtTotal" Enabled="False"
                                            Style="text-align: right" data-ToolTip="default" ToolTip="Vermelho indica que a Qtd. Embarcada está diferente da Qtd. do Destino." />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Destino:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtTotalQtdDestino" runat="server" CssClass="txtDecimal" Enabled="False"
                                            Style="text-align: right" data-ToolTip="default" ToolTip="Vermelho Indica que a Qtd. Total do Destino está diferente da Qtd. Total do Pedido." />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Pedido:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtTotalQtdPedido" runat="server" CssClass="txtNumerico9" Enabled="False"
                                            Style="text-align: right" />
                                    </div>
                                </div>
                                <div class="row" id="divRoteiroLog" runat="server">
                                    <div class="coltxt" style="width: 100%; height: 100%;">
                                        <asp:TextBox ID="txtRoteiroLog" runat="server" TextMode="MultiLine" Width="100%" Height="70px" Font-Size="6pt" ReadOnly="true" />
                                    </div>
                                </div>
                            </div>

                        </asp:Panel>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Transportes" ID="TabTransportes">
                    <HeaderTemplate>
                        Transportes
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconMais" runat="server">
                                        <asp:LinkButton ID="lnkAdicionarTransportador" Text="Adicionar Transportador" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparTransportador" Text="Limpar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Transportador:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTransportadoras" runat="server" Enabled="False" Width="600px" data-ToolTip="default" ToolTip="Informar o empresa transportadora desejada." />
                                <asp:HiddenField ID="hdfTransportadoras" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnTransportadoras" runat="server" OnClick="btnTransportadoras_Click"
                                    CssClass="btn" Text="&gt;" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data Frete:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataFrete" runat="server" CssClass="calendario" Width="86px" data-ToolTip="default" ToolTip="Data da efetivação do transporte." />
                            </div>
                            <div class="collbl">
                                Quota:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQuota" runat="server" CssClass="txtDecimal4" data-ToolTip="default" ToolTip="Limite a ser transportado." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Quantidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQuantidadeTransp" runat="server" CssClass="txtDecimal4" Width="107px" data-ToolTip="default" ToolTip="" />
                            </div>
                            <div class="collbl">
                                Unitário Frete:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtUnitarioFrete" runat="server" CssClass="txtDecimal10" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 200px;">
                            <asp:GridView ID="gridTransportes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridTransportes_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText="&gt;&gt;" ShowSelectButton="True" ButtonType="Button">
                                        <FooterStyle Width="10px"></FooterStyle>
                                        <HeaderStyle Width="10px"></HeaderStyle>
                                        <ItemStyle Width="10px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Codigo">
                                        <HeaderStyle HorizontalAlign="Left" Width="110px"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="110px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoEndereco" HeaderText="End.">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="10px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Nome">
                                        <ItemTemplate>
                                            <asp:Label ID="Label1" runat="server" Text='<%# Bind("Transportador.Nome") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="DataFrete" HeaderText="Data Frete">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="UnitarioFrete" HeaderText="Unit&#225;rio Frete">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Quantidade" DataFormatString="{0:n4}" HeaderText="Quantidade"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="QuotaDiaria" DataFormatString="{0:n4}" HeaderText="Quota"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirTransportador" runat="server" Height="18px" ImageUrl="~/Images/erro.jpg"
                                                data-ToolTip="default" ToolTip="Excluir Transportador" Width="18px" Style="cursor: pointer"
                                                OnClick="imgExcluirTransportador_Click" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Representantes" ID="TblRepresentante">
                    <HeaderTemplate>
                        Representantes
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="painelleft" style="width: 59%;">
                            <div class="menu_acoes">
                                <div class="acoes">
                                    <ul>
                                        <li class="iconMais" runat="server">
                                            <asp:LinkButton ID="lnkAdicionarRepresentante" Text="Adicionar Representante" runat="server" />
                                        </li>
                                        <li class="iconLimpar" runat="server">
                                            <asp:LinkButton ID="lnkLimparRepresentante" Text="Limpar" runat="server" />
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Representante:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtRepresentante" runat="server" Enabled="False" Width="600px" />
                                    <asp:HiddenField ID="hdfRepresentante" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnRepresentante" runat="server" OnClick="btnRepresentante_Click"
                                        Text="&gt;" CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Nome do Representante." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Tipo:
                                </div>
                                <div class="coltxt">
                                    <asp:CheckBox ID="chkRepPrincipal" runat="server" Text="Principal"
                                        data-ToolTip="default" ToolTip="Se é o principal representante da empresa, marcar a opção." />
                                </div>
                                <div class="coltxt">
                                    <asp:CheckBox ID="chkPercFixoComissao" Text="Percentual Fixo" runat="server" AutoPostBack="True"
                                        data-ToolTip="default" ToolTip="Se existe um percentual fixo ao representante, marcar a opção." />
                                </div>
                                <div class="collbl">
                                    Percentual(%) :
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPerRepre" CssClass="txtDecimal10" Width="86px" runat="server" AutoPostBack="True"
                                        data-ToolTip="default" ToolTip="Caso o percentual seja variável, inserir o valor manualmente." />
                                </div>
                                <div class="collbl">
                                    Valor:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtValorRepre" runat="server" CssClass="txtDecimal" Width="85px" BackColor="#FFFF99"
                                        Enabled="False" data-ToolTip="default" ToolTip="Calculado conforme o percentual inserido ou Valor Fixo." />
                                </div>
                            </div>
                            <div class="bordagrid" style="height: 215px;">
                                <asp:GridView ID="GridRepresentantes" runat="server" AutoGenerateColumns="False"
                                    Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="GridComissoes_SelectedIndexChanged">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText="&gt;&gt;" ShowSelectButton="True">
                                            <HeaderStyle Width="10px" />
                                            <ItemStyle Width="10px" />
                                        </asp:CommandField>
                                        <asp:BoundField DataField="CodigoRepresentante" HeaderText="Codigo">
                                            <HeaderStyle HorizontalAlign="Left" Width="110px" />
                                            <ItemStyle HorizontalAlign="Left" Width="110px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CodigoEnderecoRepresentante" HeaderText="End.">
                                            <HeaderStyle HorizontalAlign="Left" Width="10px" />
                                            <ItemStyle HorizontalAlign="Left" Width="10px" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Nome">
                                            <ItemTemplate>
                                                <asp:Label ID="Label1" runat="server" Text='<%# Bind("Representante.Nome") %>' />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Percentual" HeaderText="Percentual">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ValorComissao" DataFormatString="{0:n2}" HeaderText="Valor">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Principal">
                                            <ItemTemplate>
                                                <asp:CheckBox ID="CkPrincipal" runat="server" AutoPostBack="True" Checked='<%# Bind("Principal") %>'
                                                    Enabled="False" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="% Fixo">
                                            <ItemTemplate>
                                                <asp:CheckBox ID="chkPercentualFixo" runat="server" Checked='<%# Bind("PercentualFixo") %>'
                                                    Enabled="False" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgExcluirRepresentante" runat="server" Height="18px" ImageUrl="~/Images/erro.jpg"
                                                    OnClick="imgExcluirRepresentante_Click" Style="cursor: pointer" data-ToolTip="default"
                                                    ToolTip="Excluir Representante" Width="18px" />
                                            </ItemTemplate>
                                            <HeaderStyle Width="30px" />
                                            <ItemStyle Width="30px" />
                                        </asp:TemplateField>
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
                        <div id="divComissoes" runat="server" class="painelright" style="width: 41%;">
                            <div class="subtitulodiv">
                                <asp:Label ID="lblTabeladeComissao" runat="server" Text="Tabela de Comissão" />
                                <asp:HiddenField ID="HidLinhaRepresentante" runat="server" />
                            </div>
                            <div class="bordagrid" style="height: 215px;">
                                <asp:GridView ID="gridTabelaDeComissao" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="98%">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="NomeProduto" HeaderText="Nome" HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Tabela">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="ddlTabelaDeComissao" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlTabelaDeComissao_SelectedIndexChanged"
                                                    Width="250px" DataSource='<%# Bind("TabelasSafra") %>' DataTextField="Descricao"
                                                    DataValueField="Codigo" SelectedValue='<%# Bind("CodigoTabela") %>' />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:TemplateField>
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
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Vencimentos" ID="TabVencimentosOld">
                    <HeaderTemplate>
                        Vencimentos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="painelleft" style="width: 50%; margin-right: 4px;">
                            <div class="row">
                                <div class="collbl">
                                    Momento Financ.:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlMomentoFinanceiro" runat="server" OnSelectedIndexChanged="ddlMomentoFinanceiro_SelectedIndexChanged"
                                        AutoPostBack="True" Width="240px" Enabled="False" ToolTip="Momento Financeiro - O financeiro será gerado pela nota fiscal ou pelo pedido.">
                                        <asp:ListItem Value="0">Nenhum</asp:ListItem>
                                        <asp:ListItem Value="2">Venc. No Pedido</asp:ListItem>
                                        <asp:ListItem Value="3">Venc. Na Nota</asp:ListItem>
                                        <asp:ListItem Value="9">Virtual</asp:ListItem>
                                    </asp:DropDownList>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Condições:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="cmbCondicoes" runat="server" AutoPostBack="True" Width="415px" ToolTip="Condições de pagamento do pedido." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Data:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataVencimento" runat="server" CssClass="calendario" Width="100px" data-ToolTip="default" ToolTip="Data de vencimento das parcelas." />
                                </div>
                                <div class="collbl">
                                    Valor:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtValorVencimento" runat="server" Style="width: 100px;" CssClass="txtDecimal" data-ToolTip="default" ToolTip="Valor unitário da parcela." />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnOkVencimento" CssClass="btn" runat="server" OnClick="btnOkVencimento_Click"
                                        Text="OK" UseSubmitBehavior="False" />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 40%;">
                            <div class="row">
                                <div class="collbl">
                                    Total Parcelado:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtTotalPacelado" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal" Font-Bold="True"
                                        ForeColor="Red" Style="width: 100px;" data-ToolTip="default" ToolTip="Valor que será parcelado." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Total Pago:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtTotalPago" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal" Font-Bold="True"
                                        ForeColor="Red" Style="width: 100px;" Text="0" data-ToolTip="default" ToolTip="Valor das parcelas pagas." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Saldo:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtSaldoVencimentos" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal" Font-Bold="True"
                                        ForeColor="Red" Style="width: 100px;" Text="0" data-ToolTip="default" ToolTip="Saldo a pagar." />
                                </div>
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 225px;">
                            <asp:GridView ID="GridCondicoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="GridCondicoes_SelectedIndexChanged"
                                Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
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
                        <div id="divEntrega" runat="server" class="painelleft" style="width: 100%;">
                            <div class="row">
                                <div class="collbl" style="width: 200px;">
                                    Condição Pagam. Nota Entrega:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="lstCondicoesPgtoEntrega" runat="server" AutoPostBack="True"
                                        OnSelectedIndexChanged="lstCondicoesPgtoEntrega_SelectedIndexChanged" Width="420px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl" style="width: 200px;">
                                    Quota da Entrega:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtQuotaDeEntrega" runat="server" AutoPostBack="True" BackColor="#FFFFC0"
                                        CssClass="txtDecimal4" Font-Bold="True" ForeColor="Red" OnTextChanged="txtQuotaDeEntrega_TextChanged"
                                        Text="0" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl" style="width: 200px;">
                                    Periodicidade da Entrega:
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
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Observações" ID="TabObservacoes">
                    <HeaderTemplate>
                        Observações
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconMais" runat="server">
                                        <asp:LinkButton ID="lnkAdicionarObservacao" Text="Adicionar Observação" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Texto:
                            </div>
                            <div class="coltxt" style="width: 90.2%;">
                                <asp:TextBox ID="txtAddObservacao" runat="server" Height="40px" TextMode="MultiLine"
                                    Style="text-transform: uppercase;" Width="99%" data-ToolTip="default" ToolTip="Inserir texto caso haja uma observação a ser considerada." />
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            Observação
                        </div>
                        <div class="row">
                            <div class="coltxt lg">
                                <asp:TextBox ID="txtObservacao" runat="server" Width="99%" TextMode="MultiLine" Rows="1"
                                    Height="210px" ReadOnly="True" Style="text-transform: uppercase;" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Pedido de Troca" ID="TabPedidoTroca">
                    <HeaderTemplate>
                        Pedido de Troca
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="LnkVincular" Text="Vincular" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="LnkDesvincular" Text="Desvincular" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    Empresa:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="lblEmpresaTroca" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Cliente:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="lblClienteTroca" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Pedido:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="lblPedidoTroca" runat="server" />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft">

                            <div class="row">
                                <div class="collbl">
                                    Total Venda:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtVendaTroca" runat="server" CssClass="txtDecimal" ReadOnly="True"
                                        BackColor="#FFFF99" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Total Compra:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCompraTroca" runat="server" CssClass="txtDecimal" ReadOnly="True"
                                        BackColor="#FFFF99" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Saldo Em Aberto:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtSaldoTroca" runat="server" CssClass="txtDecimal" ReadOnly="True"
                                        BackColor="#FFFF99" />
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido(CN):
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblCNTroca" runat="server" />
                            </div>
                        </div>
                        <div class="row" id="divContaAdiantamentoTroca" runat="server">
                            <div class="collbl">
                                Conta:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlContaAdto" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlContaAdto_SelectedIndexChanged"
                                    Width="600px" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabParcelamento">
                    <HeaderTemplate>
                        Parcelamento
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div>
                            <ajaxToolkit:TabContainer ID="TabContainer3" runat="server" ActiveTabIndex="1" Width="100%" Style="clear: both;">

                                <ajaxToolkit:TabPanel runat="server" ID="TabParcelas">
                                    <HeaderTemplate>
                                        Parcelas
                                    </HeaderTemplate>
                                    <ContentTemplate>

                                        <div class="menu_acoes">
                                            <div class="acoes">
                                                <ul>
                                                    <li class="iconMais" runat="server">
                                                        <asp:LinkButton ID="LnkParcelar" Text="Parcelar" runat="server" />
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>

                                        <div>
                                            <div class="painelleft">
                                                <div class="row">
                                                    <div class="collbl">
                                                        Condições:
                                                    </div>
                                                    <div class="coltxt">
                                                        <asp:DropDownList ID="ddlCondPagPed" runat="server" AutoPostBack="True" Width="350px" />
                                                    </div>
                                                </div>
                                                <div class="row">
                                                    <div class="collbl">
                                                        Data:
                                                    </div>
                                                    <div class="coltxt">
                                                        <asp:TextBox ID="txtDataCondPagParcela" runat="server" CssClass="calendario" Width="100px" />
                                                    </div>
                                                </div>
                                            </div>

                                            <div class="painelleft">
                                                <div class="collbl row">
                                                    Total Pedido:
                                                </div>
                                                <div class="coltxt row">
                                                    <asp:TextBox ID="txtPedidoTotal" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal" Font-Bold="True" ForeColor="Red" Text="0" data-ToolTip="default" ToolTip="Valor que será parcelado." />
                                                </div>
                                            </div>

                                            <div class="painelleft">
                                                <div class="collbl row">
                                                    Total Realizado:
                                                </div>
                                                <div class="coltxt row">
                                                    <asp:TextBox ID="txtPedidoTotalPago" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal" Font-Bold="True" ForeColor="Red" Text="0" data-ToolTip="default" ToolTip="Valor das parcelas pagas." />
                                                </div>
                                            </div>

                                            <div class="painelleft">
                                                <div class="collbl row">
                                                    Saldo:
                                                </div>
                                                <div class="coltxt row">
                                                    <asp:TextBox ID="txtPedidoSaldo" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal" Font-Bold="True" ForeColor="Red" Text="0" data-ToolTip="default" ToolTip="Saldo a pagar." />
                                                </div>
                                            </div>

                                            <div class="painelleft" style="display: none;">
                                                <div class="row">
                                                    <div class="collbl">
                                                        Condição:
                                                    </div>
                                                    <div class="coltxt">
                                                        <asp:DropDownList ID="ddlCondPagNotas" runat="server" AutoPostBack="True" Width="350px" />
                                                    </div>
                                                </div>
                                                <div class="row">
                                                    <div class="collbl">
                                                        Quota da Entrega:
                                                    </div>
                                                    <div class="coltxt">
                                                        <asp:TextBox ID="txtQuotaNota" runat="server" AutoPostBack="True" BackColor="#FFFFC0" CssClass="txtDecimal4" Font-Bold="True" ForeColor="Red" Text="0" />
                                                    </div>
                                                </div>
                                                <div class="row">
                                                    <div class="collbl">
                                                        Periodicidade:
                                                    </div>
                                                    <div class="coltxt">
                                                        <asp:DropDownList ID="DropDownList4" runat="server" AutoPostBack="True" Width="350px">
                                                            <asp:ListItem Value="0">Antecipado/Fabrica/Avista</asp:ListItem>
                                                            <asp:ListItem Value="1">Na Entrega</asp:ListItem>
                                                            <asp:ListItem Value="2">Data Fixa</asp:ListItem>
                                                            <asp:ListItem Value="7">Semanal Entrega de segunda a domingo com Pagamento X Dias após</asp:ListItem>
                                                            <asp:ListItem Value="15">Quinzenal Entrega durante a quinzena com pagamento X Dias após</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>

                                        <div class="row">
                                            <div class="painelleft" style="width: 20%;">
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

                                            <div class="painelleft" style="width: 24%; margin-left: 1%">
                                                <div class="bordagrid" style="height: 215px;">
                                                    <div class="menu_acoes" style="margin-left: 1%; width: 96%;">
                                                        <div class="acoes">
                                                            <ul>
                                                                <li class="icon_confirmar" runat="server">
                                                                    <asp:LinkButton ID="LnkAtualizarParcela" Text="Atualizar" runat="server" />
                                                                </li>
                                                            </ul>
                                                        </div>
                                                    </div>

                                                    <div class="row" style="margin-left: 1%">
                                                        <div class="collbl">
                                                            Parcela:
                                                        </div>
                                                        <div class="coltxt">
                                                            <asp:TextBox ID="txtCodigoParcela" runat="server" CssClass="inteiro" Width="100px" BackColor="#FFFFC0" ReadOnly="True" />
                                                        </div>
                                                    </div>

                                                    <div class="row" style="margin-left: 1%">
                                                        <div class="collbl">
                                                            Data:
                                                        </div>
                                                        <div class="coltxt">
                                                            <asp:TextBox ID="txtDataVencParcela" runat="server" CssClass="calendario" Width="100px" />
                                                        </div>
                                                    </div>

                                                    <div class="row" style="margin-left: 1%">
                                                        <div class="collbl">
                                                            Valor:
                                                        </div>
                                                        <div class="coltxt">
                                                            <asp:TextBox ID="txtValorParcela" runat="server" CssClass="txtDecimal" Width="100px" />
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                                <ajaxToolkit:TabPanel runat="server" ID="TabResumoFinanceiro">
                                    <HeaderTemplate>
                                        Resumo Financeiro
                                    </HeaderTemplate>
                                    <ContentTemplate>
                                        <div class="painelleft bordagrid" style="width: 12%; margin-left: 0.5%; height: 225px;">


                                            <asp:GridView ID="GridOrigemFinanceiro" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                                    <asp:BoundField DataField="Origem" HeaderText="Origem">
                                                        <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
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

                                        <div class="painelleft bordagrid" style="width: 56%; margin-left: 0.5%; height: 225px;">
                                            <asp:GridView ID="gridFinanceiro" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:BoundField DataField="Tipo" HeaderText="Tipo">
                                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    </asp:BoundField>

                                                    <asp:BoundField DataField="Qualificador" HeaderText="Qualificador">
                                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    </asp:BoundField>

                                                    <asp:BoundField DataField="Situacao" HeaderText="Situacao">
                                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    </asp:BoundField>

                                                    <asp:BoundField DataField="Registro" HeaderText="ID">
                                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    </asp:BoundField>

                                                    <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    </asp:BoundField>


                                                    <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor Oficial">
                                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    </asp:BoundField>

                                                    <asp:BoundField DataField="ValorMoeda" DataFormatString="{0:N2}" HeaderText="Valor Moeda">
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


                                        <div class="painelleft bordagrid" style="width: 30%; margin-left: 0.5%; height: 240px;">
                                            <div>
                                                <div class="painelleft" style="width: 32%; margin-left: 0.5%;">
                                                    <div class="row"></div>
                                                    <div class="row collbl">ABERTOS</div>
                                                    <div class="row collbl">BAIXADOS </div>
                                                    <div class="row collbl">ADIANTAMENTO </div>
                                                    <div class="row collbl">BX-ADIANTAMENTO</div>
                                                </div>

                                                <div class="painelleft" style="width: 32%; margin-left: 0.5%;">
                                                    <div class="subtitulodiv">A PAGAR</div>
                                                    <div class="row coltxt">
                                                        <asp:TextBox ID="txtAbertoCP" CssClass="txtDecimal" runat="server" Text="0,00" BackColor="#FFFFC0" ReadOnly="True" />
                                                    </div>
                                                    <div class="row coltxt">
                                                        <asp:TextBox ID="txtBaixadoCP" CssClass="txtDecimal" runat="server" Text="0,00" BackColor="#FFFFC0" ReadOnly="True" />
                                                    </div>
                                                    <div class="row coltxt">
                                                        <asp:TextBox ID="txtAdCP" CssClass="txtDecimal" runat="server" Text="0,00" BackColor="#FFFFC0" ReadOnly="True" />
                                                    </div>
                                                    <div class="row coltxt">
                                                        <asp:TextBox ID="txtBxAdCP" CssClass="txtDecimal" runat="server" Text="0,00" BackColor="#FFFFC0" ReadOnly="True" />
                                                    </div>
                                                </div>

                                                <div class="painelleft" style="width: 32%; margin-left: 0.5%;">
                                                    <div class="subtitulodiv">A RECEBER</div>
                                                    <div class="row coltxt">
                                                        <asp:TextBox ID="txtAbertoCR" CssClass="txtDecimal" runat="server" Text="0,00" BackColor="#FFFFC0" ReadOnly="True" />
                                                    </div>
                                                    <div class="row coltxt">
                                                        <asp:TextBox ID="txtBaixadoCR" CssClass="txtDecimal" runat="server" Text="0,00" BackColor="#FFFFC0" ReadOnly="True" />
                                                    </div>
                                                    <div class="row coltxt">
                                                        <asp:TextBox ID="txtAdCR" CssClass="txtDecimal" runat="server" Text="0,00" BackColor="#FFFFC0" ReadOnly="True" />
                                                    </div>
                                                    <div class="row coltxt">
                                                        <asp:TextBox ID="txtBxAdCR" CssClass="txtDecimal" runat="server" Text="0,00" BackColor="#FFFFC0" ReadOnly="True" />
                                                    </div>
                                                </div>
                                            </div>
                                            <div class="row">
                                                <div class="collbl">VLR PEDIDO</div>
                                                <div class="collbl">REALIZADO</div>
                                                <div class="collbl">SALDO</div>
                                            </div>
                                            <div class="row">
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtResPedido" CssClass="txtDecimal" runat="server" Text="0,00" BackColor="#FFFFC0" ReadOnly="True" />
                                                </div>
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtResRealizado" CssClass="txtDecimal" runat="server" Text="0,00" BackColor="#FFFFC0" ReadOnly="True" />
                                                </div>
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtResSaldo" CssClass="txtDecimal" runat="server" Text="0,00" BackColor="#FFFFC0" ReadOnly="True" />
                                                </div>
                                            </div>
                                        </div>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                            </ajaxToolkit:TabContainer>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Contrato" ID="TabContrato">
                    <HeaderTemplate>
                        Contrato
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row" runat="server">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricaoContrato" runat="server" MaxLength="100" Width="500px" data-ToolTip="default"
                                    ToolTip="Descrição do Contrato." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Arquivo:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeDoArquivo" runat="server" Enabled="false" Style="width: 200px;" />
                                <asp:FileUpload ID="fupArquivo" OnChange="Arquivo()" runat="server" Width="120px" Font-Size="11px" ClientIDMode="Static" />
                            </div>
                        </div>
                        <div class="painelleft bordagrid" style="width: 99%; margin-left: 0.5%; height: 225px;">
                            <asp:GridView ID="gridContratos" runat="server" AutoGenerateColumns="False" ForeColor="#333333"
                                GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Codigo" HeaderText="Código">
                                        <HeaderStyle HorizontalAlign="Center" Width="70px" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                                        <HeaderStyle HorizontalAlign="Left" Width="500px" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeDoArquivo" HeaderText="Nome do arquivo">
                                        <HeaderStyle HorizontalAlign="Left" Width="300px" />
                                        <ItemStyle HorizontalAlign="Left" Width="300px" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Arquivo">
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgDownload" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/download32x32.png" Style="margin-top: 0;"
                                                Height="18px" Width="18px" OnClick="imgDownload_Click" ToolTip="Baixar Arquivo" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Excluir">
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluir" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                OnClick="imgExcluir_Click" data-ToolTip="default" ToolTip="Excluir"
                                                OnClientClick="return confirm('Deseja realmente excluir o Contrato?');" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <asp:UpdatePanel ID="updpnlAdicionar" runat="server">
                            <ContentTemplate>
                                <asp:Button ID="btnAdicionar" runat="server" Text="Adicionar" ClientIDMode="Static"
                                    OnClick="btnAdicionar_Click" CssClass="none" />
                            </ContentTemplate>
                            <Triggers>
                                <asp:PostBackTrigger ControlID="btnAdicionar" />
                            </Triggers>
                        </asp:UpdatePanel>
                        <asp:Label ID="lblMsg" runat="server" />
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Obs.Controle Interno" ID="TabControleInterno">
                    <HeaderTemplate>
                        Controle Interno
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="subtitulodiv">
                            Observação de Controle Interno
                        </div>
                        <div class="row">
                            <div class="coltxt lg">
                                <asp:TextBox ID="txtObservacaoControleInterno" runat="server" Width="99%" TextMode="MultiLine" Rows="1"
                                    Height="265px" ReadOnly="True" Style="text-transform: uppercase;" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
    <uc:ConsultaEstados ID="ucConsultaEstados" runat="server" />
    <uc:ConsultaDadosBancarios ID="ucConsultaDadosBancarios" runat="server" />
    <uc:ConsultaCodMunicipios ID="ucConsultaCodMunicipios" runat="server" />
    <uc:ConsultaPedidoDeTroca ID="ucConsultaPedidoDeTroca" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:PedidoXLancamento ID="ucPedidoXLancamento" runat="server" />
    <uc:PedidoXEncargo ID="ucPedidoxEncargo" runat="server" />
    <uc:PedidoXFixacao ID="ucPedidoxFixacao" runat="server" />
    <uc:FixacaoProcuracao ID="ucFixacaoProcuracao" runat="server" />
    <uc:ConsultaEncargosPlanoDeContas ID="ucConsultaEncargosPlanoDeContas" runat="server" />
    <uc:Supervisor ID="ucSupervisor" runat="server" />
    <uc:ConsultarNaviosXInvoice ID="ucConsultarNaviosXInvoice" runat="server" />
    <uc:EmailNFePedido ID="ucEmailNFePedido" runat="server" />
</asp:Content>
