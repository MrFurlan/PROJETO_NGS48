<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="OrdemDeCarregamento.aspx.vb" Inherits="NGS.Web.UI.OrdemDeCarregamento" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />

    <style type="text/css">
        /* Layout geral */
        #meioconteudo {
            width: 1340px !important;
        }

        /* Tabelas com borda pontilhada */
        table.pontilhado {
            width: 100%;
            border-collapse: collapse;
        }

        .pontilhado, .pontilhado td, .pontilhado th {
            border: 1px dotted #000;
            border-collapse: collapse;
        }

        /* Título e tipografia auxiliar */
        .titulo span {
            font-size: 7pt;
            text-transform: uppercase;
        }

        /* Unificação de tamanho de fonte para rótulos e campos */
        .campo {
            display: inline-block;
            min-height: 22px;
            line-height: 22px;
            vertical-align: middle;
            font-size: 7pt;
        }

        input, select, textarea {
            font-size: 7pt;
        }

        .campo-flex {
            display: flex;
            align-items: center;
            gap: 5px;
        }

        .nome {
            flex: 1;
        }

        /* Larguras típicas por conteúdo */
        .cep {
            width: 80px;
        }

        .uf {
            width: 28px;
        }

        .fone {
            width: 140px;
        }

        .ie {
            width: 140px;
        }

        .cnpj {
            width: 150px;
        }

        /* Ajustes do bloco de observação (lado direito) para calçar a altura */
        #obsCell {
            padding: 4px;
        }

        #lblObservacaoCarregamento {
            display: block;
            white-space: pre-wrap;
        }

        /* Tabelas internas: só divisórias internas, sem borda externa */
        table.inner-grid {
            width: 100%;
            border: 0;
            border-collapse: collapse;
        }

            table.inner-grid td, table.inner-grid th {
                border: 0;
                padding: 2px 4px;
            }

                table.inner-grid td + td, table.inner-grid th + th {
                    border-left: 1px dotted #000;
                }

            table.inner-grid tr + tr td, table.inner-grid tr + tr th {
                border-top: 1px dotted #000;
            }

        .td-no-pad {
            padding: 0;
        }

        .totais-cell {
            display: grid;
            grid-template-rows: auto 26px;
            row-gap: 4px;
            min-height: 54px;
        }

            .totais-cell .titulo span {
                font-size: 7pt;
                line-height: 1.1;
                white-space: normal;
                word-break: break-word;
            }

        .valor-monetario {
            display: inline-block;
            width: 100%;
            line-height: 26px;
            text-align: right;
            white-space: nowrap;
            box-sizing: border-box;
            font-size: 7pt;
        }

        /* Coluna fixa proporcional (20% da largura da tabela) */
        .col-total {
            width: 20%;
            min-width: 150px;
            box-sizing: border-box;
        }

            .col-total .valor-monetario {
                width: 100% !important;
            }

        .flex-grids {
            display: flex;
            gap: 8px;
        }

            .flex-grids .painelleft {
                flex: 1;
            }

        /* Classe utilitária para títulos de seções em Label */
        .upper {
            text-transform: uppercase;
        }
    </style>

    <script type="text/javascript">
        // Faz a Observação crescer só em altura, sem afetar largura/colunas
        function equalizarObservacao() {
            var esq = document.getElementById('tblEsq');
            var dir = document.getElementById('tblDir');
            var sizer = document.getElementById('obsSizer');
            if (!esq || !dir || !sizer) return;
            sizer.style.height = '0px';
            var diff = esq.offsetHeight - dir.offsetHeight;
            if (diff > 0) { sizer.style.height = diff + 'px'; }
        }

        function pageLoadNotaItens() {
            $("#txtDataDeEmissao:enabled").datepicker({
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

            $("#txtDataCarregamento:enabled").datepicker({
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

            $("#txtDataDeEmissao").setMask("date");
            $("#txtDataCarregamento").setMask("date");

        }

        $(document).ready(function () {
            equalizarObservacao();
            pageLoadNotaItens();
            setTimeout(equalizarObservacao, 0);
            try {
                var prm = Sys.WebForms.PageRequestManager.getInstance();
                prm.add_endRequest(function () {
                    if (typeof pageLoadNotaItens === 'function') { pageLoadNotaItens(); }
                    equalizarObservacao();
                });
            } catch (e) { }
            window.addEventListener('resize', equalizarObservacao);
        });

    </script>
</asp:Content>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngNotaFiscalXItens" runat="server" EnableScriptGlobalization="True" EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />

    <div class="titulodiv">
        ORDEM DE CARREGAMENTO
        <div class="row" style="float: right;">
            <div class="coltxt">
                <asp:Label ID="lblUsuario" runat="server" Visible="False" Font-Bold="True" Font-Size="7pt" />
            </div>
            <div class="coltxt">
                <asp:ImageButton ID="imgUsuario" runat="server" Width="18px" Height="20px" ImageAlign="AbsMiddle" ImageUrl="~/Images/man2.png" Visible="False" />
            </div>
        </div>
    </div>

    <asp:UpdatePanel ID="updpnlNotaFiscalXItens" runat="server">
        <ContentTemplate>

            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="primeiraVez" runat="server" />

            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" runat="server" Text="Gravar" /></li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkAtualizar" runat="server" Text="Atualizar" OnClientClick="if(!confirm('Deseja realmente alterar este registro?')) return false;" /></li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" runat="server" Text="Excluir" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" /></li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" runat="server" Text="Consultar" /></li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" /></li>
                        <li class="iconRelatorio rel" runat="server">
                            <a>Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton ID="lnkPdf" runat="server" CssClass="iconPdf" Text="Pdf" ToolTip="Gerar arquivo em PDF." /></li>
                                <li>
                                    <asp:LinkButton ID="lnkImpressora" runat="server" CssClass="iconRelatorio" Text="Impressora" ToolTip="Enviar Danfe para impressora." /></li>
                            </ul>
                        </li>
                        <li class="iconMail" runat="server">
                            <asp:LinkButton ID="lnkEnviarEmail" runat="server" Text="Enviar E-mail" /></li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" /></li>
                        <li runat="server">
                            <div style="float: left;">
                                <label class="titulo"><span style="font-size: 7pt; font-weight: bold; color: blue;">IP USUÁRIO:</span></label>
                                <asp:Label ID="lblAcessoUsuario" runat="server" Width="40px" Style="font-weight: bold; color: red;" />
                            </div>
                        </li>
                        <li runat="server" style="float: right;">
                            <div class="row" style="margin-top: 0;">
                                <div class="coltxt">
                                    <asp:Image ID="Image1" runat="server" Width="18px" Height="20px" ImageAlign="AbsMiddle" ImageUrl="~/Images/man2.png" ToolTip="Usuário Lançamento" />
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlUsuarios" runat="server" Width="145px" />
                                </div>
                            </div>
                        </li>
                    </ul>
                </div>
            </div>

            <!-- Container 100% com 2 colunas -->
            <table class="pontilhado" cellpadding="0" cellspacing="0" width="100%" style="margin-top: 3px;">
                <tr>
                    <!-- Esquerda -->
                    <td style="width: 50%; vertical-align: top;">
                        <table id="tblEsq" class="pontilhado" cellpadding="2" cellspacing="0" width="100%">
                            <colgroup>
                                <col style="width: 70%" />
                                <col style="width: 30%" />
                            </colgroup>

                            <!-- EMPRESA -->
                            <tr>
                                <td colspan="10"><strong>
                                    <asp:Label ID="lblEmpresa" runat="server" CssClass="upper" Text="EMPRESA" /></strong></td>
                            </tr>

                            <tr>
                                <td colspan="7" valign="top">
                                    <label class="titulo"><strong><span>NOME/RAZÃO SOCIAL</span></strong></label><br />
                                    <div class="campo-flex">
                                        <asp:Button ID="btnEmpresa" runat="server" CssClass="btn" UseSubmitBehavior="False" Text="&gt;" />
                                        <asp:Label ID="lblNomeEmpresa" runat="server" CssClass="campo nome" ForeColor="Blue" />
                                    </div>
                                </td>
                                <td colspan="3" valign="top">
                                    <label class="titulo"><strong><span>CNPJ</span></strong></label><br />
                                    <asp:Label ID="lblCnpjEmpresa" runat="server" CssClass="campo cnpj" ForeColor="Blue" />
                                </td>
                            </tr>

                            <tr>
                                <td colspan="5" valign="top">
                                    <label class="titulo"><strong><span>ENDEREÇO</span></strong></label><br />
                                    <asp:Label ID="lblEnderecoEmpresa" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="2" valign="top">
                                    <label class="titulo"><strong><span>COMPLEMENTO</span></strong></label><br />
                                    <asp:Label ID="lblComplementoEmpresa" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="2" valign="top">
                                    <label class="titulo"><strong><span>BAIRRO/DISTRITO</span></strong></label><br />
                                    <asp:Label ID="lblBairroEmpresa" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="1" valign="top">
                                    <label class="titulo"><strong><span>CEP</span></strong></label><br />
                                    <asp:Label ID="lblCepEmpresa" runat="server" CssClass="campo cep" ForeColor="Blue" />
                                </td>
                            </tr>

                            <tr>
                                <td colspan="5" valign="top">
                                    <label class="titulo"><strong><span>MUNICÍPIO</span></strong></label><br />
                                    <asp:Label ID="lblCidadeEmpresa" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="3" valign="top">
                                    <label class="titulo"><strong><span>FONE</span></strong></label><br />
                                    <asp:Label ID="lblTelefoneEmpresa" runat="server" CssClass="campo fone" ForeColor="Blue" />
                                </td>
                                <td colspan="1" valign="top">
                                    <label class="titulo"><strong><span>UF</span></strong></label><br />
                                    <asp:Label ID="lblUfEmpresa" runat="server" CssClass="campo uf" ForeColor="Blue" />
                                </td>
                                <td colspan="1" valign="top">
                                    <label class="titulo"><strong><span>INSCRIÇÃO ESTADUAL</span></strong></label><br />
                                    <asp:Label ID="lblIeEmpresa" runat="server" CssClass="campo ie" ForeColor="Blue" />
                                </td>
                            </tr>

                            <!-- TRANSPORTADOR -->
                            <tr>
                                <td colspan="10"><strong>
                                    <asp:Label ID="lblTransportador" runat="server" CssClass="upper" Text="TRANSPORTADOR E VOLUMES" /></strong></td>
                            </tr>

                            <tr>
                                <td colspan="7" valign="top">
                                    <label class="titulo"><strong><span>NOME/RAZÃO SOCIAL</span></strong></label><br />
                                    <div class="campo-flex">
                                        <asp:Button ID="btnTransportador" runat="server" CssClass="btn" UseSubmitBehavior="False" Text="&gt;" />
                                        <asp:Label ID="lblNomeTransportador" runat="server" CssClass="campo nome" ForeColor="Blue" />
                                    </div>
                                </td>
                                <td colspan="3" valign="top">
                                    <label class="titulo"><strong><span>CNPJ/CPF</span></strong></label><br />
                                    <asp:Label ID="txtCodigoTransportador" runat="server" CssClass="campo cnpj" ForeColor="Blue" />
                                </td>
                            </tr>

                            <tr>
                                <td colspan="5" valign="top">
                                    <label class="titulo"><strong><span>ENDEREÇO</span></strong></label><br />
                                    <asp:Label ID="lblEnderecoTransportador" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="2" valign="top">
                                    <label class="titulo"><strong><span>COMPLEMENTO</span></strong></label><br />
                                    <asp:Label ID="lblComplementoTransportador" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="2" valign="top">
                                    <label class="titulo"><strong><span>BAIRRO/DISTRITO</span></strong></label><br />
                                    <asp:Label ID="lblBairroTransportador" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="1" valign="top">
                                    <label class="titulo"><strong><span>CEP</span></strong></label><br />
                                    <asp:Label ID="lblCepTransportador" runat="server" CssClass="campo cep" ForeColor="Blue" />
                                </td>
                            </tr>

                            <tr>
                                <td colspan="5" valign="top">
                                    <label class="titulo"><strong><span>MUNICÍPIO</span></strong></label><br />
                                    <asp:Label ID="lblCidadeTransportador" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="3" valign="top">
                                    <label class="titulo"><strong><span>FONE/FAX</span></strong></label><br />
                                    <asp:Label ID="lblTelefoneTransportador" runat="server" CssClass="campo fone" ForeColor="Blue" />
                                </td>
                                <td colspan="1" valign="top">
                                    <label class="titulo"><strong><span>UF</span></strong></label><br />
                                    <asp:Label ID="lblUfTransportador" runat="server" CssClass="campo uf" ForeColor="Blue" />
                                </td>
                                <td colspan="1" valign="top">
                                    <label class="titulo"><strong><span>INSCRIÇÃO ESTADUAL</span></strong></label><br />
                                    <asp:Label ID="lblIeTransportador" runat="server" CssClass="campo ie" ForeColor="Blue" />
                                </td>
                            </tr>

                            <!-- MOTORISTA -->
                            <tr>
                                <td colspan="10"><strong>
                                    <asp:Label ID="lblMotorista" runat="server" CssClass="upper" Text="MOTORISTA" /></strong></td>
                            </tr>

                            <tr>
                                <td colspan="7" valign="top">
                                    <label class="titulo"><strong><span>NOME</span></strong></label><br />
                                    <div class="campo-flex">
                                        <asp:Button ID="btnMotorista" runat="server" CssClass="btn" UseSubmitBehavior="False" Text="&gt;" />
                                        <asp:Label ID="lblNomeMotorista" runat="server" CssClass="campo nome" ForeColor="Blue" />
                                    </div>
                                </td>
                                <td colspan="3" valign="top">
                                    <label class="titulo"><strong><span>CPF</span></strong></label><br />
                                    <asp:Label ID="lblCpfMotorista" runat="server" CssClass="campo cnpj" ForeColor="Blue" />
                                </td>
                            </tr>

                            <tr>
                                <td colspan="5" valign="top">
                                    <label class="titulo"><strong><span>ENDEREÇO</span></strong></label><br />
                                    <asp:Label ID="lblEnderecoMotorista" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="2" valign="top">
                                    <label class="titulo"><strong><span>COMPLEMENTO</span></strong></label><br />
                                    <asp:Label ID="lblComplementoMotorista" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="2" valign="top">
                                    <label class="titulo"><strong><span>BAIRRO/DISTRITO</span></strong></label><br />
                                    <asp:Label ID="lblBairroMotorista" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="1" valign="top">
                                    <label class="titulo"><strong><span>CEP</span></strong></label><br />
                                    <asp:Label ID="lblCepMotorista" runat="server" CssClass="campo cep" ForeColor="Blue" />
                                </td>
                            </tr>

                            <tr>
                                <td colspan="5" valign="top">
                                    <label class="titulo"><strong><span>MUNICÍPIO</span></strong></label><br />
                                    <asp:Label ID="lblCidadeMotorista" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td colspan="3" valign="top">
                                    <label class="titulo"><strong><span>FONE</span></strong></label><br />
                                    <asp:Label ID="lblTelefoneMotorista" runat="server" CssClass="campo fone" ForeColor="Blue" />
                                </td>
                                <td colspan="1" valign="top">
                                    <label class="titulo"><strong><span>UF</span></strong></label><br />
                                    <asp:Label ID="lblUfMotorista" runat="server" CssClass="campo uf" ForeColor="Blue" />
                                </td>
                                <td colspan="1" valign="top">
                                    <label class="titulo"><strong><span>CNH</span></strong></label><br />
                                    <asp:Label ID="lblCnhMotorista" runat="server" CssClass="campo ie" ForeColor="Blue" />
                                </td>
                            </tr>
                        </table>
                    </td>

                    <!-- Direita -->
                    <td style="width: 50%; vertical-align: top;">
                        <table id="tblDir" class="pontilhado" cellpadding="2" cellspacing="0" width="100%">
                            <colgroup>
                                <col style="width: 24%" />
                                <col style="width: 20%" />
                                <col style="width: 20%" />
                                <col style="width: 16%" />
                                <col style="width: 20%" />
                            </colgroup>

                            <!-- ORDEM -->
                            <tr>
                                <td colspan="5"><strong>ORDEM</strong></td>
                            </tr>
                            <tr>
                                <td valign="top">
                                    <label class="titulo"><strong><span>N° ORDEM</span></strong></label><br />
                                    <asp:Label ID="lblNumeroOrdem" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                                <td class="col-data" valign="top">
                                    <label class="titulo"><strong><span>DATA DE EMISSÃO</span></strong></label><br />
                                    <asp:Label ID="lblDataEmissao" runat="server" CssClass="campo" ForeColor="Blue" />
                                    <asp:TextBox ID="txtDataDeEmissao" runat="server" Width="60px" Enabled="True" BorderStyle="None"
                                        OnTextChanged="txtDataDeEmissao_TextChanged" AutoPostBack="True" ClientIDMode="Static" />
                                </td>
                                <td class="col-data" valign="top">
                                    <label class="titulo"><strong><span>DATA DE CARREGAMENTO</span></strong></label><br />
                                    <asp:Label ID="lblDataCarregamento" runat="server" CssClass="campo" ForeColor="Blue" />
                                    <asp:TextBox ID="txtDataCarregamento" runat="server" Width="60px" Enabled="True" BorderStyle="None"
                                        OnTextChanged="txtDataCarregamento_TextChanged" AutoPostBack="True" ClientIDMode="Static" />
                                </td>
                                <td colspan="2" valign="top">
                                    <label class="titulo"><strong><span>SITUAÇÃO</span></strong></label><br />
                                    <asp:Label ID="lblSituacaoOrdem" runat="server" CssClass="campo" ForeColor="Blue" />
                                </td>
                            </tr>

                            <!-- DADOS CAMINHÃO -->
                            <tr>
                                <td colspan="5"><strong>DADOS DO CAMINHÃO</strong></td>
                            </tr>

                            <!-- Caminhão 1 -->
                            <tr>
                                <td colspan="5" class="td-no-pad">
                                    <table class="inner-grid" cellpadding="0" cellspacing="0" width="100%">
                                        <colgroup>
                                            <col style="width: 16%" />
                                            <col style="width: 28%" />
                                            <col style="width: 6%" />
                                            <col style="width: 14%" />
                                            <col style="width: 36%" />
                                        </colgroup>
                                        <tr>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>PLACA 1</span></strong></label><br />
                                                <asp:Button ID="btnPlaca" OnClick="btnPlaca_Click" runat="server" UseSubmitBehavior="False"
                                                    Text="&gt;" CssClass="btn" />
                                                <asp:HiddenField ID="txtCodigoPlaca" runat="server" />
                                                <asp:Label ID="lblPlaca1" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>MUNICÍPIO</span></strong></label><br />
                                                <asp:Label ID="lblMunicipioPlaca1" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>UF</span></strong></label><br />
                                                <asp:Label ID="lblUfPlaca1" runat="server" CssClass="campo uf" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>RNTRC</span></strong></label><br />
                                                <asp:Label ID="lblRntrcPlaca1" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>PROPRIETÁRIO</span></strong></label><br />
                                                <asp:Label ID="lblProprietarioPlaca1" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>

                            <!-- Caminhão 2 -->
                            <tr>
                                <td colspan="5" class="td-no-pad">
                                    <table class="inner-grid" cellpadding="0" cellspacing="0" width="100%">
                                        <colgroup>
                                            <col style="width: 16%" />
                                            <col style="width: 28%" />
                                            <col style="width: 6%" />
                                            <col style="width: 14%" />
                                            <col style="width: 36%" />
                                        </colgroup>
                                        <tr>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>PLACA 2</span></strong></label><br />
                                                <asp:Label ID="lblPlaca2" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>MUNICÍPIO</span></strong></label><br />
                                                <asp:Label ID="lblMunicipioPlaca2" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>UF</span></strong></label><br />
                                                <asp:Label ID="lblUfPlaca2" runat="server" CssClass="campo uf" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>RNTRC</span></strong></label><br />
                                                <asp:Label ID="lblRntrcPlaca2" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>PROPRIETÁRIO</span></strong></label><br />
                                                <asp:Label ID="lblProprietarioPlaca2" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>

                            <!-- Caminhão 3 -->
                            <tr>
                                <td colspan="5" class="td-no-pad">
                                    <table class="inner-grid" cellpadding="0" cellspacing="0" width="100%">
                                        <colgroup>
                                            <col style="width: 16%" />
                                            <col style="width: 28%" />
                                            <col style="width: 6%" />
                                            <col style="width: 14%" />
                                            <col style="width: 36%" />
                                        </colgroup>
                                        <tr>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>PLACA 3</span></strong></label><br />
                                                <asp:Label ID="lblPlaca3" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>MUNICÍPIO</span></strong></label><br />
                                                <asp:Label ID="lblMunicipioPlaca3" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>UF</span></strong></label><br />
                                                <asp:Label ID="lblUfPlaca3" runat="server" CssClass="campo uf" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>RNTRC</span></strong></label><br />
                                                <asp:Label ID="lblRntrcPlaca3" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                            <td valign="top">
                                                <label class="titulo"><strong><span>PROPRIETÁRIO</span></strong></label><br />
                                                <asp:Label ID="lblProprietarioPlaca3" runat="server" CssClass="campo" ForeColor="Blue" /></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>

                            <!-- OBSERVAÇÃO calçando a altura -->
                            <tr>
                                <td id="obsCell" colspan="5" rowspan="4">
                                    <label class="titulo"><strong><span>OBSERVAÇÃO</span></strong></label><br />
                                    <asp:Label ID="lblObservacaoCarregamento" runat="server" CssClass="campo" ForeColor="Blue" />
                                    <div id="obsSizer" style="height: 0;"></div>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>

            <!-- O resto da página -->
            <table class="pontilhado" cellpadding="0" width="100%" style="margin-top: 3px;">
                <tr>
                    <td colspan="1" valign="top" style="white-space: nowrap;">
                        <label class="titulo"><strong><span style="font-size: 7pt;">PEDIDO</span></strong></label><br />
                        <asp:Button ID="btnPedido" runat="server" CssClass="btn" UseSubmitBehavior="False" Text="&gt;" />
                        <asp:Label ID="txtPedido" runat="server" ForeColor="Blue" Font-Size="7pt" />
                        <asp:ImageButton ID="imgExtratoPedido" runat="server" Width="16px" Height="16px" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ToolTip="Visualizar Extrato do Pedido" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo"><strong><span style="font-size: 7pt;">ROMANEIO</span></strong></label><br />
                        <asp:Button ID="btnRomaneio" runat="server" CssClass="btn" UseSubmitBehavior="False" Text="&gt;" />
                        <asp:Label ID="txtRomaneio" runat="server" ForeColor="Blue" />
                        <asp:ImageButton ID="imgRomaneio" runat="server" Width="16px" Height="16px" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ToolTip="Visualizar Romaneio" />
                    </td>
                    <td colspan="5" valign="top" style="padding: 3px; width: 50%; white-space: nowrap;">
                        <label class="titulo"><strong><span>LOCAL DE EMBARQUE/COLETA</span></strong></label><br />
                        <asp:Button ID="btnEmbarque" runat="server" CssClass="btn" UseSubmitBehavior="False" Enabled="False" OnClick="btnEmbarque_Click" Text="&gt;" />
                        <asp:DropDownList ID="ddlEmbarque" runat="server" Width="93%" Height="18px" AutoPostBack="True" OnSelectedIndexChanged="ddlEmbarque_SelectedIndexChanged"></asp:DropDownList>
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo"><strong><span>AUT.RETIRADA</span></strong></label><br />
                        <asp:Button ID="BtnRetirada" runat="server" CssClass="btn" UseSubmitBehavior="False" Text="&gt;" />
                        <asp:Label ID="txtAutorizacao" runat="server" ForeColor="Blue" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo"><strong><span>FRETE POR CONTA</span></strong></label><br />
                        <asp:DropDownList ID="ddlFrete" runat="server" Width="120px" AutoPostBack="True"></asp:DropDownList>
                    </td>
                </tr>

                <tr>
                    <td colspan="10" valign="top" style="padding: 5px;">
                        <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                            <ajaxToolkit:TabPanel ID="tbPedidos" runat="server" HeaderText="tbPedidos">
                                <HeaderTemplate>PEDIDOS</HeaderTemplate>
                                <ContentTemplate>
                                    <div class="flex-grids">
                                        <div class="painelleft" style="margin-left: 4px; margin-right: 4px;">
                                            <div class="subtitulodiv">PEDIDOS</div>
                                            <div class="bordagrid" style="height: 138px;">
                                                <asp:GridView ID="grdPedidos" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <EditRowStyle BackColor="#999999" />
                                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                    <Columns>
                                                        <asp:BoundField DataField="Codigo" HeaderText="Registro" />
                                                        <asp:BoundField DataField="ReceberPagar" HeaderText="R/P" />
                                                        <asp:BoundField DataField="DescricaoMoeda" HeaderText="Moeda" />
                                                        <asp:BoundField DataField="DescricaoProvisao" HeaderText="Situacao" />
                                                        <asp:BoundField DataField="Vencimento" HeaderText="Vencimento" DataFormatString="{0:dd/MM/yyyy}" />
                                                        <asp:BoundField DataField="ValorLiquido" HeaderText="Valor Oficial" DataFormatString="{0:n2}" />
                                                        <asp:BoundField DataField="MoedaValorLiquido" HeaderText="Valor Moeda" DataFormatString="{0:n2}" />
                                                    </Columns>
                                                </asp:GridView>
                                            </div>
                                        </div>

                                        <div class="painelleft" style="margin-left: 4px; margin-right: 4px;">
                                            <div class="subtitulodiv">ITENS PEDIDO</div>
                                            <div class="bordagrid" style="height: 138px;">
                                                <asp:GridView ID="grdItensPedido" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None">
                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                    <EditRowStyle BackColor="#999999" />
                                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                    <Columns>
                                                        <asp:BoundField DataField="Codigo" HeaderText="Registro" />
                                                        <asp:BoundField DataField="ReceberPagar" HeaderText="R/P" />
                                                        <asp:BoundField DataField="DescricaoMoeda" HeaderText="Moeda" />
                                                        <asp:BoundField DataField="DescricaoProvisao" HeaderText="Situacao" />
                                                        <asp:BoundField DataField="Vencimento" HeaderText="Vencimento" DataFormatString="{0:dd/MM/yyyy}" />
                                                        <asp:BoundField DataField="ValorLiquido" HeaderText="Valor Oficial" DataFormatString="{0:n2}" />
                                                        <asp:BoundField DataField="MoedaValorLiquido" HeaderText="Valor Moeda" DataFormatString="{0:n2}" />
                                                    </Columns>
                                                </asp:GridView>
                                            </div>
                                        </div>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>

                            <ajaxToolkit:TabPanel ID="tbNotas" runat="server" HeaderText="tbNotas">
                                <HeaderTemplate>NOTAS</HeaderTemplate>
                                <ContentTemplate>
                                    <div class="painelleft" style="margin-left: 4px; margin-right: 4px; width: 100%;">
                                        <div class="subtitulodiv">NOTAS</div>
                                        <div class="bordagrid" style="height: 138px;">
                                            <asp:GridView ID="grdNotas" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <EditRowStyle BackColor="#999999" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="#FFFFFF" HorizontalAlign="Center" />
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <Columns>
                                                    <asp:BoundField DataField="Codigo" HeaderText="Registro" />
                                                    <asp:BoundField DataField="ReceberPagar" HeaderText="R/P" />
                                                    <asp:BoundField DataField="DescricaoMoeda" HeaderText="Moeda" />
                                                    <asp:BoundField DataField="DescricaoProvisao" HeaderText="Situacao" />
                                                    <asp:BoundField DataField="Vencimento" HeaderText="Vencimento" DataFormatString="{0:dd/MM/yyyy}" />
                                                    <asp:BoundField DataField="ValorLiquido" HeaderText="Valor Oficial" DataFormatString="{0:n2}" />
                                                    <asp:BoundField DataField="MoedaValorLiquido" HeaderText="Valor Moeda" DataFormatString="{0:n2}" />
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                        </ajaxToolkit:TabContainer>
                    </td>
                </tr>

                <!-- Totais / volumes / observações -->
                <tr>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>BASE DE CÁLCULO DO ICMS</span></strong></label>
                        <asp:TextBox ID="txtBaseIcmsNota" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>VALOR DO ICMS</span></strong></label>
                        <asp:TextBox ID="txtValorIcmsNota" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>BASE DE CÁLCULO DO ICMS SUBSTITUIÇÃO</span></strong></label>
                        <asp:TextBox ID="txtValorBaseIcmsSTNota" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>VALOR DO ICMS SUBSTITUIÇÃO</span></strong></label>
                        <asp:TextBox ID="txtValorIcmsSTNota" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td colspan="4"></td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>VALOR TOTAL DOS PRODUTOS</span></strong></label>
                        <asp:TextBox ID="txtValorTotalDosProdutos" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                </tr>
                <tr>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>VALOR DO FRETE</span></strong></label><br />
                        <asp:TextBox ID="txtValorFrete" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>VALOR DO SEGURO</span></strong></label><br />
                        <asp:TextBox ID="txtSeguro" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>DESCONTO</span></strong></label><br />
                        <asp:TextBox ID="txtDesconto" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>OUTRAS DESPESAS ACESSÓRIAS</span></strong></label><br />
                        <asp:TextBox ID="txtOutras" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>VALOR DO IPI</span></strong></label><br />
                        <asp:TextBox ID="txtValorIPINota" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td colspan="3"></td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>VALOR TOTAL DA NOTA</span></strong></label><br />
                        <asp:TextBox ID="txtValorTotalDaNota" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                </tr>
                <tr style="height: 32px;">
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>QUANTIDADE DE VOLUME(S)</span></strong></label><br />
                        <asp:TextBox ID="txtVolumes" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo"><strong><span>ESPÉCIE</span></strong></label><br />
                        <asp:TextBox ID="txtEspecie" runat="server" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo"><strong><span>MARCA</span></strong></label><br />
                        <asp:TextBox ID="txtMarca" runat="server" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>NUMERAÇÃO</span></strong></label><br />
                        <asp:TextBox ID="txtNumeracao" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td></td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>PESO ROMANEIO</span></strong></label><br />
                        <asp:TextBox ID="txtPesoRomaneio" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>PESO BRUTO</span></strong></label><br />
                        <asp:TextBox ID="txtPesoBruto" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                    <td class="col-total" valign="top">
                        <label class="titulo"><strong><span>PESO LÍQUIDO</span></strong></label><br />
                        <asp:TextBox ID="txtPesoLiquido" runat="server" CssClass="txtDecimal valor-monetario" Enabled="False" BorderColor="White" ForeColor="Blue" />
                    </td>
                </tr>
            </table>

            <table class="pontilhado" cellpadding="0" width="100%" style="margin-top: 3px;">
                <tr>
                    <td colspan="10">
                        <table style="border: none; width: 100%;">
                            <tr>
                                <td valign="top" style="width: 33%; border: medium none;"><strong>DADOS ADICIONAIS</strong></td>
                                <td valign="top" style="width: 34%; border: medium none;"><strong>OBSERVAÇÕES FISCAIS</strong><asp:Button ID="btnObservacoesFiscais" runat="server" CssClass="btn" UseSubmitBehavior="False" OnClick="btnObservacoesFiscais_Click" Text="&gt;" /></td>
                                <td valign="top" style="width: 33%; border: medium none;"><strong>OBSERVAÇÕES CONTROLE INTERNO</strong></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="10">
                        <table style="border: none; width: 100%;">
                            <tr>
                                <td align="center" valign="top" style="width: 33%; border: medium none;">
                                    <asp:TextBox ID="txtObservacoesDeEmbarque" runat="server" Width="98%" Height="155px" TextMode="MultiLine" ForeColor="Blue" OnTextChanged="txtObservacoesDeEmbarque_TextChanged" />
                                </td>
                                <td align="center" valign="top" style="width: 34%; border: medium none;">
                                    <asp:TextBox ID="txtObservacoesFiscais" runat="server" ClientIDMode="Static" Width="98%" Height="155px" TextMode="MultiLine" ForeColor="Blue" OnTextChanged="txtObservacoesFiscais_TextChanged" />
                                </td>
                                <td align="center" valign="top" style="width: 33%; border: medium none;">
                                    <asp:TextBox ID="txtObservacoesInternas" runat="server" Width="98%" Height="155px" TextMode="MultiLine" ForeColor="Blue" AutoPostBack="True" OnTextChanged="txtObservacoesInternas_TextChanged" />
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>

    <!-- UserControls -->
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:PedidoXSaldo ID="ucPedidoxSaldo" runat="server" />
    <uc:ConsultaNotaTroca ID="ucConsultaNotaTroca" runat="server" />
    <uc:ConsultaRomaneios ID="ucConsultaRomaneios" runat="server" />
    <uc:ConsultaObservacoes ID="ucConsultaObservacoes" runat="server" />
    <uc:ConsultaPlacas ID="ucConsultaPlacas" runat="server" />
    <uc:ConsultaEstados ID="ucConsultaEstados" runat="server" />
    <uc:ConsultaCodMunicipios ID="ucConsultaCodMunicipios" runat="server" />
    <uc:NotaFiscalXClassificacao ID="ucNotaFiscalXClassificacao" runat="server" />
    <uc:ConsultaNotaVendaAOrdem ID="ucConsultaNotaVendaAOrdem" runat="server" />
    <uc:ConsultaAutorizacaoDeRetirada ID="ucConsultaAutorizacaoDeRetirada" runat="server" />
    <uc:ConsultaEncargosPlanoDeContas ID="ucConsultaEncargosPlanoDeContas" runat="server" />
    <uc:ConsultaDadosBancarios ID="ucConsultaDadosBancarios" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
    <uc:NotaFiscalReferencial ID="ucNotaFiscalReferencial" runat="server" />
    <uc:NotaDeDevolucaoXNota ID="ucNotaDeDevolucaoXNota" runat="server" />
    <uc:ConsultaProcuracao ID="ucConsultaProcuracao" runat="server" />
    <uc:Inutilizacao ID="ucInutilizacao" runat="server" />
    <uc:Vencimentos ID="ucVencimentos" runat="server" />
    <uc:MonitorDeNotas ID="ucMonitorDeNotas" runat="server" />
    <uc:EmailNFe ID="ucEmailNFe" runat="server" />
    <uc:NFObsProduto ID="ucNFObsProduto" runat="server" />
    <uc:NFEncargo ID="ucNFEncargo" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
    <uc:NFProdutor ID="ucNFProdutor" runat="server" />
    <uc:NFReferencialSaida ID="ucNFReferencialSaida" runat="server" />
    <uc:ConsultaDeLote ID="ucConsultaLote" runat="server" />
    <uc:ConsultarNaviosXInvoice ID="ucConsultarNaviosXInvoice" runat="server" />
    <uc:ConsultarTransferencias ID="ucTransferencias" runat="server" />
</asp:Content>
