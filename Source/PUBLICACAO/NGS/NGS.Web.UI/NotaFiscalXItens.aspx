<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="NotaFiscalXItens.aspx.vb" Inherits="NGS.Web.UI.NotaFiscalXItens" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        #meioconteudo {
            width: 1340px !important;
        }

        table.pontilhado {
            border-collapse: collapse;
            width: 100%;
        }

            table.pontilhado > tbody > tr > td {
                border: 1px dotted #000000;
                padding: 3px;
            }

        .collbl {
            font-size: 8pt;
            width: 136px;
        }
    </style>
    <script type="text/javascript">
        function pageLoadNotaItens() {
            $("#txtDataDeEntrada:enabled").datepicker({
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

            $("#txtDataDeEmissao").setMask("date");
            $("#txtDataDeEntrada").setMask("date");

            $("#txtChaveNFe").live('keypress', function (e) {
                var code = (e.keyCode ? e.keyCode : e.which);
                if (code == 13) {
                    $("#MainContent_imgVerificarChaveNFE").click();
                    return false;
                }
            });

            $("#fupArquivo").focus(function () {
                return false;
            });
        }

        $(document).ready(function () {
            pageLoadNotaItens();
            var prmNotaItens = Sys.WebForms.PageRequestManager.getInstance();
            prmNotaItens.add_endRequest(pageLoadNotaItens);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngNotaFiscalXItens" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:ajaxupdating id="ajaxUpdating" runat="server" text="Aguarde..." />
    <div class="titulodiv">
        Nota Fiscal
        <div class="row" style="float: right;">
            <div class="coltxt">
                <asp:Label ID="lblUsuario" runat="server" Font-Bold="True" Visible="False" Font-Size="11px" />
            </div>
            <div class="coltxt">
                <asp:ImageButton ID="imgUsuario" runat="server" Width="18px" Height="20px" ImageAlign="AbsMiddle"
                    ImageUrl="~/Images/man2.png" Visible="False" />
            </div>
        </div>
    </div>
    <asp:UpdatePanel ID="updpnlNotaFiscalXItens" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="primeiraVez" runat="server" />
            <asp:HiddenField ID="observacaoValores" runat="server" />
            <asp:HiddenField ID="observacaoAutorizacao" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" OnClientClick="if(!confirm('Deseja realmente alterar este registro?')) return false;" />
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkEspelho" Text="Espelho" runat="server" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkServicoSefaz" Text="Serviço Sefaz" runat="server" data-ToolTip="default"
                                ToolTip="Verifica na Sefaz se o Serviço(NFe) está em operação" />
                        </li>
                        <li class="iconRelatorio rel" runat="server"><a>DANFE</a>
                            <ul>
                                <li>
                                    <asp:LinkButton class="iconPdf" ID="lnkPdf" runat="server" Text="Pdf" data-ToolTip="default" ToolTip="Gerar arquivo em PDF." />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconRelatorio" ID="lnkImpressora" runat="server" Text="Impressora" data-ToolTip="default" ToolTip="Enviar Danfe para impressora." />
                                </li>
                            </ul>
                        </li>
                        <li class="iconMail" runat="server">
                            <asp:LinkButton ID="lnkEnviarSEFAZ" Text="Enviar SEFAZ" runat="server" />
                        </li>
                        <li class="iconMail" runat="server">
                            <asp:LinkButton ID="lnkEnviarEmail" Text="Enviar E-mail" runat="server" />
                        </li>
                        <li class="iconRelatorio rel" runat="server" id="IdVisualizarNFe" visible="false"><a>Visualizar NFe</a>
                            <ul>
                                <li runat="server" id="IdModelNota" visible="false">
                                    <asp:LinkButton class="iconPdf" ID="lkn_ModeloNota" Text="PDF" runat="server" />
                                </li>
                                <li runat="server" id="IdVisualizar" visible="false">
                                    <asp:UpdatePanel ID="updpnlVisualizar" runat="server">
                                        <ContentTemplate>
                                            <asp:LinkButton class="iconRelatorio" ID="lnkVisualizar" Text="Txt do Xml" runat="server" />
                                        </ContentTemplate>
                                        <Triggers>
                                            <asp:PostBackTrigger ControlID="lnkVisualizar" />
                                        </Triggers>
                                    </asp:UpdatePanel>
                                </li>
                            </ul>
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                        <li runat="server">
                            <div class="row" style="margin-top: 0;">
                                <div class="coltxt">
                                    <asp:Button ID="btnModo" CssClass="btn" runat="server" BackColor="Green" BorderStyle="None"
                                        Enabled="False" Font-Bold="True" ForeColor="White" Style="cursor: pointer; font-family: Tahoma,Arial,Helvetica,sans-serif; font-size: 11px; height: 24px; text-align: center; width: 150px;"
                                        Text="MODO NORMAL" UseSubmitBehavior="False" />
                                </div>
                            </div>
                        </li>
                        <li>
                            <div class="coltxt">
                                <asp:CheckBox ID="chkImportarProdutoUnico" ClientIDMode="Static" runat="server" Text="Importar produto único" AutoPostBack="True" OnCheckedChanged="chkImportarXML_CheckedChanged"
                                    Font-Bold="True" data-ToolTip="default" ToolTip="Importar produto único" GroupName="IMPORTARXML" />
                            </div>
                        </li>
                        <li>
                            <div class="coltxt">
                                <asp:CheckBox ID="chkAGruparNCM" ClientIDMode="Static" runat="server" Text="Agrupar NCM" AutoPostBack="True" OnCheckedChanged="chkImportarXML_CheckedChanged"
                                    Font-Bold="True" data-ToolTip="default" ToolTip="Agrupar NCM" GroupName="IMPORTARXML" />
                            </div>
                        </li>
                        <li runat="server">
                            <label class="titulo">
                                <span style="font-size: 6pt; font-weight: bold; color: blue">IP USUÁRIO:</span></label>
                            <asp:Label ID="lblAcessoUsuario" runat="server" Width="40px" Style="font-weight: bold; color: red" />
                        </li>
                        <li runat="server" style="float: right;">
                            <div class="row" style="margin-top: 0;">
                                <div class="coltxt">
                                    <asp:Image ID="Image1" runat="server" Height="20px" ImageAlign="AbsMiddle" ImageUrl="~/Images/man2.png"
                                        data-ToolTip="default" ToolTip="Usuário Lançamento" Width="18px" />
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlUsuarios" runat="server" Width="145px">
                                    </asp:DropDownList>
                                </div>
                            </div>
                        </li>
                    </ul>
                </div>
            </div>
            <table class="pontilhado" cellpadding="0" width="100%" style="margin-top: 3px;">
                <tr>
                    <td colspan="3" valign="top">
                        <label class="titulo">
                            <span style="font-size: 6pt;"><strong>NATUREZA DA OPERAÇÃO</strong></span>
                        </label>
                        <br />
                        <asp:Label ID="txtNaturezaDaOperacao" runat="server" BorderStyle="None" Font-Bold="False"
                            ForeColor="Blue" />
                    </td>
                    <td colspan="3" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 8px;">CHAVE DE ACESSO DO MDF-e CONSULTA NO SITE: </span><span style="font-size: 9px;"><a target="_blank" href="http://WWW.NFE.FAZENDA.GOV.BR">WWW.NFE.FAZENDA.GOV.BR</a></span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtChaveNFe" CssClass="chaveNFe" ClientIDMode="Static" runat="server"
                            Width="320px" BorderColor="White" ReadOnly="True" TabIndex="2" />
                        <asp:LinkButton ID="lnkVerificarChaveNFE" CssClass="lnk"
                            data-tooltip="default" ToolTip="Consultar/Validar Chave NFE" runat="server" Text=" &gt; "
                            OnClick="lnkVerificarChaveNFE_Click"> <i class="fa fa-arrow-right seta"></i></asp:LinkButton>
                        <br />
                        <label class="titulo" id="txtTextoSegCodBarra" runat="server" visible="false">
                            <strong><span style="font-size: 8px;">SEGUNDO CODIGO DE BARRAS:</span></strong>
                        </label>
                        <asp:TextBox ID="txtSegCodBarra" CssClass="txtNumerico" ClientIDMode="Static" runat="server"
                            Width="320px" BorderColor="White" MaxLength="36" ReadOnly="True" Visible="false" />
                    </td>
                    <td colspan="2" valign="middle">
                        <table>
                            <tr id="trFile" runat="server" visible="true">
                                <td style="width: 47px; white-space: nowrap; padding: 0;" valign="top">
                                    <span class="HeaderSpanSecond">Arquivo:</span>
                                </td>
                                <td>
                                    <uc:file id="ucFile" runat="server" />
                                </td>
                            </tr>
                        </table>
                    </td>
                    <td colspan="1" valign="middle" align="center">
                        <asp:CheckBox ID="chkEspelho" TabIndex="3" runat="server" Text="Imp. Espelho" />
                    </td>
                    <td valign="middle" style="font-size: 7pt;" align="center">
                        <asp:Button ID="BtnRecontabilizar" runat="server" CssClass="botao" OnClick="BtnRecontabilizar_Click"
                            Style="margin: 0px;" Text="Recontabilizar" UseSubmitBehavior="False" />
                    </td>
                </tr>
                <tr>
                    <td colspan="2" valign="top" style="white-space: nowrap;">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">NOME/RAZÃO SOCIAL</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="btnEmpresa" OnClick="btnEmpresa_Click" runat="server" UseSubmitBehavior="False"
                            Text="&gt;" CssClass="btn" />
                        <asp:Label ID="txtNomeDaEmpresa" runat="server" Font-Bold="False" ForeColor="Blue"
                            Font-Italic="False" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtInscricaoDaEmpresa" runat="server" ForeColor="Blue" Font-Italic="False"
                            Font-Names="Tahoma" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">INS. ESTADUAL DO SUBST. TRIBUTARIO</span></strong>
                        </label>
                        <br />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">CNPJ</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtCnpjDaEmpresa" runat="server" ForeColor="Blue" Font-Italic="False"
                            Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="middle" style="font-size: 7pt;" align="center">
                        <strong>NÚMERO:</strong>
                    </td>
                    <td colspan="1" valign="middle" style="font-size: 7pt;" align="center">
                        <asp:TextBox ID="txtNumero" runat="server" Width="72px" ForeColor="Blue" OnTextChanged="txtNumero_TextChanged"
                            BorderColor="White" Font-Names="Tahoma" Style="text-align: right" CssClass="txtNumerico9"
                            AutoPostBack="True" ToolTip="Número de registro referente a nota fiscal." />
                    </td>
                    <td colspan="1" valign="middle" align="center">
                        <table>
                            <tr>
                                <td valign="middle" style="font-size: 7pt;" align="center">1-SAÍDA
                                    <br />
                                    2-ENTRADA
                                </td>
                                <td>
                                    <asp:Label ID="txtES" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="10pt"
                                        ForeColor="Blue" />
                                </td>
                            </tr>
                        </table>
                    </td>
                    <td colspan="1" valign="middle" align="center">
                        <asp:Button ID="BtnTransferencia" runat="server" CssClass="botao" OnClick="BtnTransferencia_Click"
                            Style="margin: 0px;" Text="Transferências" UseSubmitBehavior="False" />
                    </td>
                </tr>
                <tr>
                    <td colspan="6">
                        <strong>
                            <asp:Label ID="lblCliente" runat="server" Text="Label" /></strong>
                    </td>
                    <td colspan="4">&nbsp;
                    </td>
                </tr>
                <tr>
                    <td colspan="4" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">NOME/RAZÃO SOCIAL</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="btnCliente" OnClick="btnCliente_Click" runat="server" UseSubmitBehavior="False"
                            Text="&gt;" CssClass="btn" />
                        <asp:Label ID="txtNomeDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                        <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">CNPJ/CPF</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtCnpjDoCliente" runat="server" ForeColor="Blue" Font-Italic="False"
                            Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">DATA DA EMISSÃO</span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtDataDeEmissao" runat="server" Width="60px" Enabled="True" BorderStyle="None"
                            OnTextChanged="txtDataDeEmissao_TextChanged" AutoPostBack="True" ClientIDMode="Static" />
                    </td>
                    <td colspan="1" style="font-size: 7pt;" valign="middle" align="center">
                        <strong>SÉRIE:</strong>
                    </td>
                    <td colspan="1" style="font-size: 7pt;" valign="middle" align="center">
                        <asp:TextBox ID="txtSerie" runat="server" Width="72px" ForeColor="Blue" OnTextChanged="txtSerie_TextChanged"
                            MaxLength="3" BorderColor="White" Style="text-align: right" AutoPostBack="True" />
                    </td>
                    <td style="font-size: 7pt; white-space: nowrap;" align="center" valign="middle">
                        <strong>
                            <asp:CheckBox ID="chk_nfe" runat="server" Checked="false" AutoPostBack="True" Enabled="False" OnCheckedChanged="chk_nfe_CheckedChanged"
                                Text="Nota Eletrônica" />
                        </strong>
                    </td>
                    <td colspan="1" align="center" valign="middle">
                        <asp:Button ID="btnMonitorDeNotas" OnClick="btnMonitorDeNotas_Click" runat="server"
                            CssClass="botao" UseSubmitBehavior="False" Text="Monitor NF" Style="margin: 0px;"></asp:Button>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">ENDEREÇO</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtEnderecoDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtComplementoDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">BAIRRO/DISTRITO</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtBairroDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">CEP</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtCepDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">DATA SAÍDA/ENTRADA</span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtDataDeEntrada" runat="server" Width="60px" Enabled="True" BorderStyle="None"
                            OnTextChanged="txtDataDeEntrada_TextChanged" AutoPostBack="True" ClientIDMode="Static" />
                    </td>
                    <td colspan="1" valign="middle" style="font-size: 7pt;" align="center">
                        <strong>NUM. VIAS</strong>
                    </td>
                    <td colspan="1" valign="middle" style="font-size: 7pt;" align="center">
                        <asp:TextBox ID="txtNumViasDeImpressao" runat="server" CssClass="txtInteiro" Width="72px"
                            ForeColor="Blue" MaxLength="2" BorderColor="White" Font-Names="Tahoma" Style="text-align: right;" />
                    </td>
                    <td style="font-size: 7pt; white-space: nowrap;" align="center" valign="middle">
                        <strong>
                            <asp:CheckBox ID="chk_NossaEmissao" runat="server" Enabled="False" Text="Nossa Emissão"
                                AutoPostBack="True" OnCheckedChanged="chk_NossaEmissao_CheckedChanged" />
                        </strong>
                    </td>
                    <td style="font-size: 7pt; white-space: nowrap;" align="center" valign="middle">
                        <asp:Button ID="btnInutilizar" runat="server" CssClass="botao" OnClick="btnInutilizar_Click"
                            Style="margin: 0px;" Text="Inutilizar" UseSubmitBehavior="False" />
                        <asp:Button ID="btnReajusta" runat="server" Text="Reajusta" />
                        <asp:TextBox ID="txtReajuste" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">MUNICÍPIO</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtCidadeDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">FONE/FAX</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtTelefoneDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">UF</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtEstadoDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtInscricaoDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt;">HORA DA SAÍDA</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtHoraDaSaida" runat="server" Font-Bold="False" ForeColor="Blue"
                            Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top" style="font-size: 6pt;">
                        <strong>SITUAÇÃO</strong>
                        <br />
                        <asp:TextBox ID="txtSituacao" runat="server" Width="60px" Enabled="False" MaxLength="50"
                            BorderColor="White" data-ToolTip="default" ToolTip="Situação Nota Fiscal" />
                    </td>
                    <td colspan="3" valign="top" style="font-size: 6pt;">
                        <strong>TIPO DOCUMENTO</strong>
                        <br />
                        <asp:DropDownList ID="ddlTipoDeDocumento" runat="server" Width="98%" Enabled="False"
                            AutoPostBack="True" OnSelectedIndexChanged="ddlTipoDeDocumento_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td colspan="1" valign="top" style="white-space: nowrap;">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">PEDIDO</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="btnPedido" OnClick="btnPedido_Click" runat="server" UseSubmitBehavior="False"
                            Text="&gt;" CssClass="btn" />
                        <asp:Label ID="txtPedido" runat="server" ForeColor="Blue" Font-Names="Tahoma" Font-Size="7pt" />
                        <asp:ImageButton ID="imgExtratoPedido" OnClick="imgExtratoPedido_Click" runat="server"
                            Width="16px" Height="16px" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                            data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">OPERAÇÃO COMERCIAL</span></strong>
                        </label>
                        <br />
                        <asp:DropDownList ID="cmbSubOperacao" runat="server" Width="100%" Enabled="False"
                            AutoPostBack="True">
                        </asp:DropDownList>
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">ROMANEIO</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="btnRomaneio" OnClick="btnRomaneio_Click" runat="server" UseSubmitBehavior="False"
                            Text="&gt;" CssClass="btn" />
                        <asp:Label ID="txtRomaneio" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                        <asp:ImageButton ID="imgRomaneio" runat="server" Width="16px" Height="16px" ImageAlign="AbsMiddle"
                            ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" data-ToolTip="default"
                            ToolTip="Visualizar Romaneio" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">CLASSIF.</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="btnClassificacao" OnClick="btnClassificacao_Click" runat="server"
                            UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">AUT.RETIRADA</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="BtnRetirada" OnClick="btnRetirada_Click" runat="server" UseSubmitBehavior="False"
                            Text="&gt;" CssClass="btn" />
                        <asp:Label ID="txtAutorizacao" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">Cessão de Crédito</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="btnProcuracao" OnClick="btnProcuracao_Click" runat="server" UseSubmitBehavior="False"
                            Text="&gt;" CssClass="btn" />
                        <asp:Label ID="txtCessaoDeCredito" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">TROCA DE NOTA</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="BtnNotas" runat="server" OnClick="BtnNotas_Click" Text="Notas" UseSubmitBehavior="False"
                            Width="100%" />
                        <asp:CheckBox ID="ChkTroca" runat="server" Text="Sim" Visible="False" AutoPostBack="True" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">FINALIDADE</span></strong>
                        </label>
                        <br />
                        <asp:DropDownList ID="cmbFinalidade" runat="server" Width="98%" AutoPostBack="True"
                            OnSelectedIndexChanged="cmbFinalidade_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td colspan="5" valign="top" style="padding: 3px; width: 50%; white-space: nowrap;">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt;">DEPÓSITO</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="btnDeposito" OnClick="btnDeposito_Click" runat="server" Text="&gt;"
                            CssClass="btn" UseSubmitBehavior="False" Enabled="False" />
                        <asp:DropDownList ID="ddlDeposito" runat="server" Width="93%" Height="18px" AutoPostBack="True"
                            OnSelectedIndexChanged="ddlDeposito_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                    <td colspan="5" valign="top" style="padding: 3px; width: 50%; white-space: nowrap;">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt;">LOCAL DE EMBARQUE/COLETA</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="btnEmbarque" OnClick="btnEmbarque_Click" runat="server" Text="&gt;"
                            CssClass="btn" UseSubmitBehavior="False" Enabled="False" />
                        <asp:DropDownList ID="ddlEmbarque" runat="server" Width="93%" Height="18px" AutoPostBack="True"
                            OnSelectedIndexChanged="ddlEmbarque_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td colspan="5" valign="top" style="padding: 3px; width: 50%; white-space: nowrap;">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt;">ORIGEM / DESTINO</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="BtnOrigemDestino" OnClick="BtnOrigemDestino_Click" runat="server"
                            Text="&gt;" CssClass="btn" UseSubmitBehavior="False" Enabled="False" />
                        <asp:DropDownList ID="ddlOrigemDestino" runat="server" Width="93%" Height="18px"
                            AutoPostBack="True" OnSelectedIndexChanged="ddlOrigemDestino_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                    <td colspan="5" valign="top" style="padding: 3px; width: 50%; white-space: nowrap;">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt;">TRANSBORDO</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="BtnTransbordo" OnClick="BtnTransbordo_Click" runat="server" Text="&gt;"
                            CssClass="btn" UseSubmitBehavior="False" Enabled="False" />
                        <asp:DropDownList ID="ddlTransbordo" runat="server" Width="93%" Height="18px" AutoPostBack="True"
                            OnSelectedIndexChanged="ddlTransbordo_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td colspan="5" valign="top" style="padding: 3px; width: 50%; white-space: nowrap;">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt;">ENTREGA</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="BtnEntrega" OnClick="BtnEntrega_Click" runat="server"
                            Text="&gt;" CssClass="btn" UseSubmitBehavior="False" Enabled="False" />
                        <asp:DropDownList ID="ddlEntrega" runat="server" Width="93%" Height="18px"
                            AutoPostBack="True" OnSelectedIndexChanged="ddlEntrega_SelectedIndexChanged">
                        </asp:DropDownList>
                    </td>
                    <td colspan="5" valign="top" style="padding: 3px; width: 50%; white-space: nowrap;">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt;"></span></strong>
                        </label>
                        <br />
                        <asp:Button ID="BtnVazio" runat="server" Text="&gt;"
                            CssClass="btn" UseSubmitBehavior="False" Enabled="False" />
                        <asp:DropDownList ID="ddlVazio" runat="server" Width="93%" Height="18px" Enabled="False">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr id="divNaviosXInvoice" runat="server" visible="false">
                    <td colspan="5" valign="top" style="padding: 3px; width: 50%; white-space: nowrap;">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt;">INVOICE</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="bntConsultarNaviosXInvoice" runat="server" Text="&gt;" OnClick="bntConsultarNaviosXInvoice_Click" CssClass="btn" UseSubmitBehavior="False" />
                        <asp:Label ID="txtNaviosXInvoice" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="5" valign="top" style="padding: 3px; width: 50%; white-space: nowrap;"></td>
                </tr>
                <tr>
                    <td colspan="10" valign="top" style="padding: 5px;">
                        <ajaxtoolkit:tabcontainer id="TabContainer1" runat="server" activetabindex="1" width="100%">
                            <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel1" ID="TabPanel1">
                                <HeaderTemplate>
                                    Produtos / Serviços
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="bordagrid" style="height: 200px;">
                                        <asp:GridView ID="gridItens" runat="server" Width="100%" ForeColor="#333333" OnSelectedIndexChanged="gridItens_SelectedIndexChanged" OnRowCreated="gridItens_RowCreated" GridLines="None"
                                            CellPadding="4" AutoGenerateColumns="False">
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                                            <Columns>
                                                <asp:CommandField InsertText="" SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                                                <asp:BoundField DataField="Produto" HeaderText="Produto">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="NomeProduto" HeaderText="Nome Produto">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Sequencia" HeaderText="Seq.">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Lote" HeaderText="Lote">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Classificacao" HeaderText="Class.">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Embalagem" HeaderText="Emb.">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:TemplateField HeaderText="Saldo">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Literal ID="LtrSaldo" runat="server" Text='<%# Eval("Saldo", "{0:N4}")%>' />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Qtd. Física">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Literal ID="LtrQuantidadeFisica" runat="server" Text='<%# Eval("QuantidadeFisica", "{0:N4}") %>' />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Quantidade">
                                                    <HeaderStyle HorizontalAlign="Right" />
                                                    <ItemStyle HorizontalAlign="Right" />
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtQuantidadeItem" runat="server" CssClass="txtDecimal4" Text='<%# Eval("Quantidade", "{0:N4}") %>' />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="BaseCalculo" HeaderText="Base">
                                                    <HeaderStyle HorizontalAlign="Right" />
                                                    <ItemStyle HorizontalAlign="Right" />
                                                </asp:BoundField>
                                                <asp:TemplateField HeaderText="Unitário">
                                                    <HeaderStyle HorizontalAlign="Right" />
                                                    <ItemStyle HorizontalAlign="Right" />
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtUnitarioItem" CssClass="txtDecimal10" runat="server" Text='<%# Eval("Unitario", "{0:N10}") %>' />&nbsp;
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Total">
                                                    <HeaderStyle HorizontalAlign="Right" />
                                                    <ItemStyle HorizontalAlign="Right" />
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtTotalItem" runat="server" Enabled="False" CssClass="txtDecimal"
                                                            Text='<%# Eval("Total", "{0:N2}") %>' />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Encargos" ShowHeader="False">
                                                    <HeaderStyle HorizontalAlign="Center" />
                                                    <ItemTemplate>
                                                        <asp:Button ID="btnEncargos" runat="server" Text=" + " UseSubmitBehavior="False"
                                                            OnClick="btnEncargos_Click" />
                                                    </ItemTemplate>
                                                    <ItemStyle Width="30px" HorizontalAlign="Center" />
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Obs.">
                                                    <HeaderStyle HorizontalAlign="Center" />
                                                    <ItemStyle HorizontalAlign="Center" />
                                                    <ItemTemplate>
                                                        <asp:ImageButton ID="imgObsProduto" runat="server"
                                                            Width="18px" Height="18px" ImageUrl="~/images/ico_OBS_ativo.gif" data-ToolTip="default"
                                                            ToolTip="Observação do produto, Msg. de Devolução e Controle de Peças." OnClick="imgObsProduto_Click" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Dev.">
                                                    <HeaderStyle HorizontalAlign="Center" />
                                                    <ItemStyle HorizontalAlign="Center" />
                                                    <ItemTemplate>
                                                        <asp:ImageButton ID="imgSelecionar" OnClick="imgSelecionar_Click" runat="server"
                                                            Width="18px" Height="18px" ImageUrl="~/images/icone-devolucao.png" data-ToolTip="default"
                                                            ToolTip="Notas Fiscais Devolvidas ou Notas Fiscais Referenciais de Exportação" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:TemplateField>
                                                    <HeaderStyle HorizontalAlign="Center" />
                                                    <ItemStyle HorizontalAlign="Center" />
                                                    <ItemTemplate>
                                                        <asp:Button ID="btnItem" runat="server" UseSubmitBehavior="False" Text="OK" CommandName="OK" OnClick="btnItem_Click"></asp:Button>
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                            </Columns>
                                            <EditRowStyle BackColor="#999999"></EditRowStyle>
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></FooterStyle>
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></HeaderStyle>
                                            <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White"></PagerStyle>
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333"></RowStyle>
                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>
                                        </asp:GridView>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel runat="server" HeaderText="TabVencimentosOld" ID="TabVencimentosOld">
                                <HeaderTemplate>
                                    Vencimentos
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="painelleft" style="margin-left: 4px; margin-right: 4px; width: 49.5%;">
                                        <div class="subtitulodiv">
                                            Pedido
                                        </div>
                                        <div class="bordagrid" style="height: 138px;">
                                            <asp:GridView AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None"
                                                ID="gridVencimentosPedido" runat="server" Width="100%">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <EditRowStyle BackColor="#999999" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <Columns>
<%--                                                    <asp:BoundField HeaderText="Registro" DataField="Codigo">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>--%>
                                                    <asp:TemplateField HeaderText="Registro">
                                                        <ItemTemplate>
                                                            <asp:HyperLink style="display:none;" ID="hpTitulo" runat="server"  NavigateUrl="#" Text='<%# Eval("Codigo")%>' />
                                                            <a href="#" onclick="newTab('<%# Eval("CodigoCifrado", "WFTitulo.aspx?param={0}")%>')"><%# Eval("Codigo")%></a>
                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="ReceberPagar" HeaderText="R/P" />
                                                    <asp:BoundField DataField="DescricaoMoeda" HeaderText="Moeda" />
                                                    <asp:BoundField DataField="DescricaoProvisao" HeaderText="Situacao" />
                                                    <asp:BoundField DataField="Prorrogacao" HeaderText="Vencimento" DataFormatString="{0:dd/MM/yyyy}">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="ValorDoDocumento" DataFormatString="{0:n2}" HeaderText="Valor Oficial">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="MoedaValorDoDocumento" DataFormatString="{0:n2}" HeaderText="Valor Moeda">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                    </div>
                                    <div class="painelleft" style="width: 49.5%;">
                                        <div class="subtitulodiv">
                                            Nota
                                        </div>
                                        <div class="bordagrid" style="height: 138px;">
                                            <asp:GridView ID="gridVencimentosNota" runat="server" Width="100%" ForeColor="#333333"
                                                GridLines="None" CellPadding="4" AutoGenerateColumns="False">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <EditRowStyle BackColor="#999999" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <Columns>
<%--                                                    <asp:BoundField DataField="Codigo" HeaderText="Registro">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>--%>
                                                    <asp:TemplateField HeaderText="Registro">
                                                        <ItemTemplate>
                                                            <asp:HyperLink style="display:none;" ID="hpTitulo" runat="server"  NavigateUrl="#" Text='<%# Eval("Codigo")%>' />
                                                            <a href="#" onclick="newTab('<%# Eval("CodigoCifrado", "WFTitulo.aspx?param={0}")%>')"><%# Eval("Codigo")%></a>
                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="ReceberPagar" HeaderText="R/P">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="DescricaoMoeda" HeaderText="Moeda">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="DescricaoProvisao" HeaderText="Situacao">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Prorrogacao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="ValorDoDocumento" DataFormatString="{0:n2}" HeaderText="Valor Oficial">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="MoedaValorDoDocumento" DataFormatString="{0:n2}" HeaderText="Valor Moeda">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                        <div class="row">
                                            <div class="coltxt">
                                                <asp:Button ID="BtnVencimentos" runat="server" CssClass="botao" Text="Atualizar"
                                                    UseSubmitBehavior="False" OnClick="BtnVencimentos_Click" />
                                            </div>
                                        </div>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel5" ID="TabPanel5">
                                <HeaderTemplate>
                                    Lote / Classificação / Embalagem
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="bordagrid" style="height: 200px;">
                                        <asp:GridView ID="gridEmbalagem" runat="server" Width="100%" Height="88px" ForeColor="#333333"
                                            GridLines="None" CellPadding="4" AutoGenerateColumns="False">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333"></RowStyle>
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></FooterStyle>
                                            <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White"></PagerStyle>
                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></HeaderStyle>
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                                            <EditRowStyle BackColor="#999999"></EditRowStyle>
                                            <Columns>
                                                <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                                                <asp:TemplateField HeaderText="Qtde Embalagem">
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="TxtQtdeEmbalagem" runat="server" Width="81px" AutoPostBack="True"
                                                            OnTextChanged="TxtQtdeEmbalagem_TextChanged" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="DescEmbalagem" HeaderText="Embalagem"></asp:BoundField>
                                                <asp:BoundField DataField="DescTipoDeEmbalagem" HeaderText="Tipo"></asp:BoundField>
                                                <asp:BoundField DataField="CapacidadeEmbalagem" HeaderText="Capacidade"></asp:BoundField>
                                                <asp:BoundField DataField="Lote" HeaderText="Lote"></asp:BoundField>
                                                <asp:BoundField DataField="DataDeValidade" HeaderText="Validade" DataFormatString="{0:dd/MM/yyyy}"></asp:BoundField>
                                                <asp:BoundField DataField="Classificacao" HeaderText="Classificação"></asp:BoundField>
                                                <asp:BoundField DataField="PesoSaco" DataFormatString="{0:N2}" HeaderText="Peso Saco"></asp:BoundField>
                                                <asp:TemplateField HeaderText="Qtde Produto">
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtQtdeDeProduto" runat="server" Enabled="False" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="Unidade" HeaderText="Un."></asp:BoundField>
                                                <asp:BoundField DataField="SaldoEmbalagem" HeaderText="Saldo Emb."></asp:BoundField>
                                                <asp:BoundField DataField="SaldoFisico" HeaderText="Etq.Fisico"></asp:BoundField>
                                                <asp:BoundField DataField="SaldoFiscal" HeaderText="Etq.Fiscal"></asp:BoundField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                    <div class="row">
                                        <asp:Button ID="BtnEmbalagem" OnClick="BtnEmbalagem_Click" runat="server" Text="Confirmar"
                                            CssClass="botao" UseSubmitBehavior="False" />
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel6" ID="TabPanel6">
                                <HeaderTemplate>
                                    Obs. Cancelamento
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="row">
                                        <asp:TextBox ID="txtObservacaoCancelamento" runat="server" Height="194px" TextMode="MultiLine"
                                            Width="100%" />
                                    </div>
                                    <div class="row">
                                        <asp:Button ID="BtnCancelarSefaz" runat="server" OnClick="BtnCancelarSefaz_Click"
                                            CssClass="botao" UseSubmitBehavior="False" Text="Cancelar Sefaz" />
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel7" ID="TabPanel7">
                                <HeaderTemplate>
                                    Contabilização
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="bordagrid" style="height: 200px;">
                                        <asp:GridView ID="gridContabilizacao" runat="server" Width="100%" ForeColor="#333333" GridLines="None"
                                            CellPadding="4" AutoGenerateColumns="False">
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></FooterStyle>
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333"></RowStyle>
                                            <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White"></PagerStyle>
                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></HeaderStyle>
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                                            <EditRowStyle BackColor="#999999"></EditRowStyle>
                                            <Columns>
                                                <asp:BoundField DataField="CodigoConta" HeaderText="Conta">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:TemplateField HeaderText="Cliente">
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblCliente" runat="server" Text='<%# Eval("CodigoCliente") & " - " & Eval("EnderecoCliente")%>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Descrição">
                                                    <ItemTemplate>
                                                        <asp:Label ID="lblDescricao" runat="server" Text='<%# Eval("Conta.Titulo")%>' />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>

                                                <asp:BoundField DataField="Movimento" HeaderText="Movimento" HtmlEncode="False" DataFormatString="{0:dd/MM/yyyy}">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Lote" HeaderText="Lote">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="CodigoProduto" HeaderText="Produto">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="CodigoCusto" HeaderText="Custo">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Historico" HeaderText="Histórico">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="DebitoOficial" DataFormatString="{0:N2}" HeaderText="D&#233;bito">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="CreditoOficial" DataFormatString="{0:N2}" HeaderText="Cr&#233;dito">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Saldo" DataFormatString="{0:N2}" HeaderText="Saldo">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>


                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>

                            <ajaxToolkit:TabPanel runat="server" HeaderText="Importação" ID="TabImportacao">
                                <%--<HeaderTemplate>
                                    Importação
                                </HeaderTemplate>--%>
                                <ContentTemplate>
                                    <asp:Panel ID="divImportacao" runat="server" CssClass="bordagrid" Style="padding: 4px;"
                                        Width="99%" Height="230px">
                                        <div class="row">
                                            <div class="collbl" style="width: 160px;">
                                                DECLARAÇÃO DE IMPORTAÇÃO:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtDI" runat="server" CssClass="txtNumerico11" />
                                            </div>
                                            <div class="collbl">
                                                DATA DECL. DE IMP.:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtDataDI" CssClass="calendario" runat="server" Width="77px" />
                                            </div>
                                            <div class="collbl">
                                                LOCAL DE EMBARQUE:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtEmbarqueDI" runat="server" Width="400px" MaxLength="60" />
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="collbl" style="width: 160px;">
                                                REGISTRO DE IMPORTAÇÃO:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="TxtRegImportacao" runat="server" />
                                            </div>
                                            <div class="collbl">
                                                DATA REG. IMP.:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="TxtDataRegImp" CssClass="calendario" runat="server" Width="77px" />
                                            </div>
                                            <div class="collbl">
                                                UF DE EMBARQUE:
                                            </div>
                                            <div class="coltxt">
                                                <asp:Button ID="btnUFEmbarqueDI" OnClick="btnUFEmbarqueDI_Click" runat="server" Text=">"
                                                    CssClass="btn" UseSubmitBehavior="False" />
                                            </div>
                                            <div class="coltxt" style="width: 141px;">
                                                <asp:Label ID="lblUFEmbarqueDI" runat="server" Style="font-size: 6pt; font-weight: bold;" />
                                                <asp:HiddenField ID="txtCodigoUFEmbarqueDI" runat="server" />
                                                &nbsp;
                                            </div>
                                            <div class="collbl">
                                                DATA EMBARQUE:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtDataEmbarqueDI" CssClass="calendario" runat="server" Width="77px" />
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="collbl" style="width: 160px;">
                                                NUM. ATO CONCESSÓRIO:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="TxtNumAtoConcessorioImp" runat="server" />
                                            </div>
                                            <div class="collbl">
                                                DATA REGISTRO:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="TxtDtaRegAtoConcessorioImp" CssClass="calendario" runat="server"
                                                    Width="77px" />
                                            </div>
                                            <div class="collbl">
                                                LOCAL DE DESEMBARQUE:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtDesembarqueDI" runat="server" Width="400px" MaxLength="60" />
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="collbl" style="width: 160px;">
                                                NR FATURA:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="TxtNrFaturaImp" runat="server" />
                                            </div>
                                            <div class="collbl">
                                                DATA VALIDADE:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="TxtDtaValidAtoConcessorioImp" CssClass="calendario" runat="server"
                                                    Width="77px" />
                                            </div>
                                            <div class="collbl">
                                                UF DE DESEMBARQUE:
                                            </div>
                                            <div class="coltxt">
                                                <asp:Button ID="btnUFDesembarqueDI" OnClick="btnUFDesembarqueDI_Click" runat="server"
                                                    CssClass="btn" Text=">" UseSubmitBehavior="False" />
                                            </div>
                                            <div class="coltxt" style="width: 141px;">
                                                <asp:Label ID="lblUFDesembarqueDI" runat="server" Style="font-size: 6pt; font-weight: bold;" />
                                                <asp:HiddenField ID="txtCodigoUFDesembarqueDI" runat="server" />
                                                &nbsp;
                                            </div>
                                            <div class="collbl">
                                                DATA DESEMBARQUE:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtDataDesembarqueDI" CssClass="calendario" runat="server" Width="77px" />
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="collbl" style="width: 160px;">
                                                VIA DE TRANSPORTE:
                                            </div>
                                            <div class="coltxt">
                                                <asp:DropDownList ID="ddlViaDeTransporteDI" runat="server" Width="370px" />
                                            </div>
                                            <div class="collbl">
                                                TIPO DE IMPORTAÇÃO
                                            </div>
                                            <div class="coltxt">
                                                <asp:DropDownList ID="ddlTipoDeImportacaoDI" runat="server" Width="410px">
                                                    <asp:ListItem Value="">Selecione</asp:ListItem>
                                                    <asp:ListItem Value="1">1 - Importação por conta própria</asp:ListItem>
                                                    <asp:ListItem Value="2">2 - Importação por conta e ordem</asp:ListItem>
                                                    <asp:ListItem Value="3">3 - Importação por encomenda</asp:ListItem>
                                                </asp:DropDownList>
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="collbl" style="width: 160px;">
                                                Valor VAFRMM:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtValorVAFRMMDI" runat="server" CssClass="txtDecimal"></asp:TextBox>
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="collbl" style="width: 160px;">
                                                FABRICANTE:
                                            </div>
                                            <div class="coltxt">
                                                <asp:Button ID="btnFabricanteDI" OnClick="btnFabricanteDI_Click" runat="server" Text=">"
                                                    CssClass="btn" UseSubmitBehavior="False" />
                                            </div>
                                            <div class="coltxt">
                                                <asp:Label ID="lblFabricanteDI" runat="server" />
                                                <asp:HiddenField ID="txtCodigolblFabricanteDI" runat="server" />
                                            </div>
                                            <div class="coltxt">
                                                <asp:ImageButton ID="imgConfirmarDI" OnClick="imgConfirmarDI_Click" runat="server"
                                                    Enabled="False" ImageUrl="~/images/confirmar.gif" data-ToolTip="default" ToolTip="Confirmar informações da DI" />
                                            </div>
                                        </div>
                                    </asp:Panel>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel runat="server" HeaderText="Exportação" ID="TabExportacao">
                                <%--<HeaderTemplate>
                                    Exportação
                                </HeaderTemplate>--%>
                                <ContentTemplate>
                                    <asp:Panel ID="pnlExportacao" runat="server" Width="100%" CssClass="borda">
                                        <table width="100%">
                                            <tr>
                                                <td valign="top" align="left" rowspan="2">
                                                    <table class="borda" width="100%" style="border: 0;">
                                                        <tr>
                                                            <td class="rotulo">Nr. Fatura de Exportação:
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtFaturaExportacao" MaxLength="10" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Nr. Despacho:
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtNrDespacho" CssClass="txtNumerico11" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Data:
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtDataDespacho" CssClass="calendario" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Navio:
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtNavio" runat="server" Width="290px" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Pais Destino:
                                                            </td>
                                                            <td>
                                                                <asp:DropDownList ID="ddlPaisDestino" runat="server" Width="292px">
                                                                </asp:DropDownList>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Dta Averbação Declar. Export.:
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtDataAverba" CssClass="calendario" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Num. Ato Concessório (Drawbak):
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="TxtNumAtoConcessorio" runat="server" CssClass="txtNumerico11" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Data Registro:
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="TxtDtaRegAtoConcessorio" CssClass="calendario" runat="server" Height="22px" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Data Validade:
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="TxtDtaValidAtoConcessorio" CssClass="calendario" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo" align="right" colspan="2">&nbsp;<asp:Button ID="BtnExportacaoSalvar" OnClick="BtnExportacaoSalvar_Click" runat="server"
                                                                CssClass="botao" Text="Salvar" UseSubmitBehavior="False" />
                                                                &nbsp;&nbsp;<asp:Button ID="BtnLimparExportacao" OnClick="BtnLimparExportacao_Click"
                                                                    runat="server" CssClass="botao" Text="Limpar Exportação" UseSubmitBehavior="False"
                                                                    Width="130px" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="2">Notas Referenciais
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="2">
                                                                <asp:GridView runat="server" AutoGenerateColumns="False" CellPadding="4" GridLines="None"
                                                                    ForeColor="#333333" Width="96%" ID="grdNotasReferenciais" ShowFooter="True">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                                                                    <Columns>
                                                                        <asp:BoundField DataField="Nota_id" HeaderText="Nota">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="Serie_id" HeaderText="Serie"></asp:BoundField>
                                                                        <asp:BoundField DataField="Quantidade" DataFormatString="{0:n2}" HeaderText="Qtd.">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="Valor" DataFormatString="{0:n2}" HeaderText="Valor">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <asp:TemplateField>
                                                                            <ItemTemplate>
                                                                                <asp:ImageButton ID="imgDelNFReferencial" runat="server" Height="20px" ImageUrl="~/images/erro.jpg"
                                                                                    Width="20px" OnClick="imgDelNFReferencial_Click" OnClientClick="if(!(confirm('Confirma Exclusão da Nota Referencial?'))) return false;" />
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                    </Columns>
                                                                    <EditRowStyle BackColor="#999999"></EditRowStyle>
                                                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></FooterStyle>
                                                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></HeaderStyle>
                                                                    <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White"></PagerStyle>
                                                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333"></RowStyle>
                                                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>
                                                                </asp:GridView>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo" align="right" colspan="2">
                                                                <asp:Button ID="btnSalvarNotasReferenciais" runat="server" CssClass="botao" Text="Salvar"
                                                                    UseSubmitBehavior="False" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td valign="top" align="left">
                                                    <table class="borda" width="100%">
                                                        <tr>
                                                            <td class="rotulo">Nr. Conhecim.:
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtNrConhecimento" runat="server" />
                                                            </td>
                                                            <td class="rotulo">Data:
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtDataNrConhecim" CssClass="calendario" runat="server" Width="70px" />
                                                            </td>
                                                            <td class="rotulo">Tipo Conhec.:
                                                            </td>
                                                            <td>
                                                                <asp:DropDownList ID="ddlTipoConhec" runat="server" Width="144px">
                                                                    <asp:ListItem Value="01">AWB</asp:ListItem>
                                                                    <asp:ListItem Value="02">MAWB</asp:ListItem>
                                                                    <asp:ListItem Value="03">HAWB</asp:ListItem>
                                                                    <asp:ListItem Value="04">COMAT</asp:ListItem>
                                                                    <asp:ListItem Value="06">R. EXPRESSAS</asp:ListItem>
                                                                    <asp:ListItem Value="07">ETIQ. REXPRESSAS</asp:ListItem>
                                                                    <asp:ListItem Value="08">HR. EXPRESSAS</asp:ListItem>
                                                                    <asp:ListItem Value="09">AV7</asp:ListItem>
                                                                    <asp:ListItem Value="10">BL</asp:ListItem>
                                                                    <asp:ListItem Value="11">MBL</asp:ListItem>
                                                                    <asp:ListItem Value="12">HBL</asp:ListItem>
                                                                    <asp:ListItem Value="13">CRT</asp:ListItem>
                                                                    <asp:ListItem Value="14">DSIC</asp:ListItem>
                                                                    <asp:ListItem Value="16">COMAT BL</asp:ListItem>
                                                                    <asp:ListItem Value="17">RWB</asp:ListItem>
                                                                    <asp:ListItem Value="18">HRWB</asp:ListItem>
                                                                    <asp:ListItem Value="19">TIF/DTA</asp:ListItem>
                                                                    <asp:ListItem Value="20">CP2</asp:ListItem>
                                                                    <asp:ListItem Value="91">N&#194;O IATA</asp:ListItem>
                                                                    <asp:ListItem Value="92">MNAO IATA</asp:ListItem>
                                                                    <asp:ListItem Value="93">HNAO IATA</asp:ListItem>
                                                                    <asp:ListItem Value="99">OUTROS</asp:ListItem>
                                                                </asp:DropDownList>
                                                            </td>
                                                            <td>
                                                                <asp:ImageButton ID="ImgAddConhec" OnClick="ImgAddConhec_Click" runat="server" ImageUrl="~/images/detalhes.png"
                                                                    ImageAlign="AbsMiddle" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="7">
                                                                <asp:Panel ID="Panel7" runat="server" Width="99%" Height="190px" CssClass="borda"
                                                                    ScrollBars="Vertical">
                                                                    <asp:GridView runat="server" AutoGenerateColumns="False" CellPadding="5" GridLines="None"
                                                                        PageSize="5" ForeColor="#333333" Width="97%" ID="GridConhecimento">
                                                                        <AlternatingRowStyle BackColor="White"></AlternatingRowStyle>
                                                                        <Columns>
                                                                            <asp:BoundField DataField="ConhecimentoDeEmbarque" HeaderText="Num. Conhecimento"></asp:BoundField>
                                                                            <asp:BoundField DataField="DataConhecimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Conhecimento"></asp:BoundField>
                                                                            <asp:BoundField DataField="TipoConhecimento" HeaderText="Cod Tipo Conhecimento" Visible="False"></asp:BoundField>
                                                                            <asp:BoundField DataField="DescTipoConhecimento" HeaderText="Tipo Conhecimento"></asp:BoundField>
                                                                            <asp:TemplateField>
                                                                                <ItemTemplate>
                                                                                    <asp:ImageButton ID="ImgExcluirConhec" runat="server" Height="20px" ImageAlign="Middle"
                                                                                        ImageUrl="~/images/erro.jpg" Width="20px" OnClick="ImgExcluirConhec_Click" />
                                                                                </ItemTemplate>
                                                                            </asp:TemplateField>
                                                                        </Columns>
                                                                        <EditRowStyle BackColor="#2461BF"></EditRowStyle>
                                                                        <EmptyDataTemplate>
                                                                        </EmptyDataTemplate>
                                                                        <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White"></FooterStyle>
                                                                        <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White"></HeaderStyle>
                                                                        <PagerStyle HorizontalAlign="Center" BackColor="#2461BF" ForeColor="White"></PagerStyle>
                                                                        <RowStyle BackColor="#EFF3FB"></RowStyle>
                                                                        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>
                                                                    </asp:GridView>
                                                                </asp:Panel>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td valign="top" align="left">
                                                    <table class="borda" width="100%">
                                                        <tr>
                                                            <td class="rotulo">Reg.Exp:
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtRe" runat="server" />
                                                            </td>
                                                            <td class="rotulo">Data:
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtDataRe" CssClass="calendario" runat="server" Width="70px" />
                                                            </td>
                                                            <td class="rotulo">UF Prod/Fab.
                                                            </td>
                                                            <td>
                                                                <asp:DropDownList ID="ddlUF" runat="server" Width="55px">
                                                                </asp:DropDownList>
                                                            </td>
                                                            <td style="width: 37px">
                                                                <asp:ImageButton ID="imgAdicionarRE" OnClick="imgAdicionarRE_Click" runat="server"
                                                                    ImageUrl="~/images/detalhes.png" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td colspan="7">
                                                                <asp:Panel ID="pnlRe" runat="server" Width="99%" Height="190px" CssClass="borda"
                                                                    ScrollBars="Vertical">
                                                                    <asp:GridView runat="server" AutoGenerateColumns="False" CellPadding="4" GridLines="None"
                                                                        ForeColor="#333333" Width="96%" ID="gridRE">
                                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                                                                        <Columns>
                                                                            <asp:BoundField DataField="RegistrodeExportacao" HeaderText="R.E."></asp:BoundField>
                                                                            <asp:BoundField DataField="DataRegistroDeExportacao" DataFormatString="{0:dd/MM/yyyy}"
                                                                                HeaderText="Data R.E."></asp:BoundField>
                                                                            <asp:BoundField DataField="UFProdutor" HeaderText="UF"></asp:BoundField>
                                                                            <asp:TemplateField>
                                                                                <ItemTemplate>
                                                                                    <asp:ImageButton ID="imgDeletarRe" runat="server" Height="20px" ImageUrl="~/images/erro.jpg"
                                                                                        Width="20px" OnClick="imgDeletarRe_Click" />
                                                                                </ItemTemplate>
                                                                            </asp:TemplateField>
                                                                        </Columns>
                                                                        <EditRowStyle BackColor="#999999"></EditRowStyle>
                                                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></FooterStyle>
                                                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></HeaderStyle>
                                                                        <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White"></PagerStyle>
                                                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333"></RowStyle>
                                                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>
                                                                    </asp:GridView>
                                                                </asp:Panel>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabVencimentos" runat="server" HeaderText="TabVencimentos">
                                <HeaderTemplate>
                                    Vencimentos
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <uc:Financeiro ID="ucFinanceiro" runat="server" />
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabNotasReferenciadas" runat="server" HeaderText="TabNotasReferenciadas">
                                <HeaderTemplate>
                                    Notas Referenciadas
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="bordagrid" style="height: 215px;">
                                        <asp:GridView ID="grdNotasReferenciadas" runat="server" AutoGenerateColumns="False"
                                            CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EditRowStyle BackColor="#999999" />
                                            <Columns>
                                                <asp:BoundField DataField="Nota_id" HeaderText="Nota"></asp:BoundField>
                                                <asp:BoundField DataField="TipoDeDocumento" HeaderText="Tipo de Documento" />
                                                <asp:BoundField DataField="Empresa_Id" HeaderText="Empresa"></asp:BoundField>
                                                <asp:BoundField DataField="EndEmpresa_Id" HeaderText="EndEmpresa"></asp:BoundField>
                                                <asp:BoundField DataField="Cliente_Id" HeaderText="Cliente"></asp:BoundField>
                                                <asp:BoundField DataField="EndCliente_Id" HeaderText="EndCliente"></asp:BoundField>
                                                <asp:BoundField DataField="ClienteNome" HeaderText="Nome"></asp:BoundField>
                                                <asp:BoundField DataField="EntradaSaida_Id" HeaderText="EntradaSaida"></asp:BoundField>
                                                <asp:BoundField DataField="Serie_Id" HeaderText="Serie"></asp:BoundField>
                                                <asp:BoundField DataField="Produto_Id" HeaderText="Produto"></asp:BoundField>
                                                <asp:BoundField DataField="CFOP_Id" HeaderText="CFOP"></asp:BoundField>
                                                <asp:BoundField DataField="Sequencia_Id" HeaderText="Sequencia"></asp:BoundField>
                                                <asp:BoundField DataField="Quantidade" HeaderText="Quantidade"></asp:BoundField>
                                                <asp:BoundField DataField="Valor" HeaderText="Valor"></asp:BoundField>
                                                <asp:BoundField DataField="TipoReferencial_Id" HeaderText="TipoReferencial"></asp:BoundField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                    <div class="row">
                                        <asp:Button ID="bntBuscarNFReferencial" CssClass="botao" runat="server" Text="Buscar" Visible="false" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Buscar Nota para Complementar" />
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabNotasDeProdutor" runat="server" HeaderText="TabNotasDeProdutor">
                                <HeaderTemplate>
                                    Notas de Produtor
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="row">
                                        <div class="bordagrid" style="height: 215px;">
                                            <asp:GridView ID="gridNotasDeProdutor" runat="server" AutoGenerateColumns="False"
                                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <EditRowStyle BackColor="#999999" />
                                                <Columns>
                                                    <asp:BoundField DataField="Data" HeaderText="Data" DataFormatString="{0:dd/MM/yyyy}">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Operacao" HeaderText="Operação">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Serie" HeaderText="Série">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Nota" HeaderText="Nota">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" DataFormatString="{0:n4}" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Valor" HeaderText="Valor" DataFormatString="{0:n2}" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <asp:Button ID="bntBuscarNFProdutor" CssClass="botao" runat="server" Text="Buscar" Visible="false" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Buscar Nota do Produtor" />
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                        </ajaxtoolkit:tabcontainer>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">BASE DE CÁLCULO DO ICMS</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtBaseIcmsNota" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">VALOR DO ICMS</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtValorIcmsNota" runat="server" Font-Bold="False" ForeColor="Blue"
                            Font-Italic="False" Font-Names="Arial" />
                    </td>
                    <td colspan="3" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">BASE DE CÁLCULO DO ICMS SUBSTITUIÇÃO</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtValorBaseIcmsSTNota" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">VALOR DO ICMS SUBSTITUIÇÃO</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtValorIcmsSTNota" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">VALOR TOTAL DOS PRODUTOS</span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtValorTotalDosProdutos" Width="50%" CssClass="txtDecimal" runat="server"
                            Enabled="False" ForeColor="Blue" BorderColor="White" Font-Names="Tahoma" />
                    </td>
                </tr>
                <tr>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">VALOR DO FRETE</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtValorFrete" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">VALOR DO SEGURO</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtSeguro" runat="server" ForeColor="Blue" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">DESCONTO</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtDesconto" runat="server" ForeColor="Blue" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">OUTRAS DESPESAS ACESSÓRIAS</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtOutras" runat="server" ForeColor="Blue" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">VALOR DO IPI</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtValorIPINota" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">VALOR TOTAL DA NOTA</span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtValorTotalDaNota" CssClass="txtDecimal" runat="server" Enabled="False"
                            ForeColor="Blue" Width="50%" BorderColor="White" Font-Names="Tahoma" />
                    </td>
                </tr>
                <tr>
                    <td colspan="10" valign="middle">
                        <strong>TRANSPORTADORA E VOLUMES IMPORTADOS</strong>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">RAZÃO SOCIAL</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="btnTransportador" OnClick="btnTransportador_Click" runat="server"
                            UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />
                        <asp:HiddenField ID="txtCodigoTransportador" runat="server" />
                        <asp:Label ID="txtNomeDoTransportador" runat="server" Width="260px" Font-Bold="False"
                            ForeColor="Blue" Font-Names="Tahoma" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">FRETE POR CONTA</span></strong>
                        </label>
                        <br />
                        <asp:DropDownList ID="ddlFrete" runat="server" AutoPostBack="True" Width="120px">
                        </asp:DropDownList>
                    </td>
                    <td colspan="1" valign="top">
                        <label>
                            <strong><span style="font-size: 6pt">CÓDIGO ANTT</span></strong>
                        </label>
                        <span style="font-size: 5pt">
                            <br />
                        </span>
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">PLACA/S DO VEÍCULO/S</span></strong>
                        </label>
                        <br />
                        <asp:Button ID="btnPlaca" OnClick="btnPlaca_Click" runat="server" UseSubmitBehavior="False"
                            Text="&gt;" CssClass="btn" />
                        <asp:HiddenField ID="txtCodigoPlaca" runat="server" />
                        <asp:Label ID="txtPlacas" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                    <td colspan="1" valign="top">
                        <span style="font-size: 6pt"><strong>UF</strong></span>
                        <br />
                        <asp:Label ID="txtEstadoDaPlaca" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                </tr>
                <tr style="height: 35px;">
                    <td colspan="3" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">ENDEREÇO</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtEnderecoDoTransportador" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">MUNICÍPIO</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtCidadeDoTransportador" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">UF</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtEstadoDoTransportador" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">CNPJ/CPF</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtCnpjDoTransportador" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong>
                        </label>
                        <br />
                        <asp:Label ID="txtInscricaoDoTransportador" runat="server" Font-Bold="False" ForeColor="Blue" />
                    </td>
                </tr>
                <tr style="height: 32px;">
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">QUANTIDADE DE VOLUME(S)</span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtVolumes" runat="server" CssClass="txtDecimal4" ForeColor="Blue"
                            BorderColor="White" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">ESPÉCIE</span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtEspecie" runat="server" ForeColor="Blue" BorderColor="White"
                            Font-Names="Tahoma" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">MARCA</span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtMarca" runat="server" ForeColor="Blue" BorderColor="White" Font-Names="Tahoma" />
                    </td>
                    <td colspan="2" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">NUMERAÇÃO</span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtNumeracao" runat="server" Width="300px" ForeColor="Blue" BorderColor="White"
                            Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">PESO ROMANEIO</span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtPesoRomaneio" runat="server" CssClass="txtNumerico" Enabled="False"
                            ForeColor="Blue" BorderColor="White" Font-Names="Tahoma" Width="65px" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">PESO BRUTO</span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtPesoBruto" runat="server" CssClass="txtDecimal3" ForeColor="Blue"
                            BorderColor="White" Font-Names="Tahoma" />
                    </td>
                    <td colspan="1" valign="top">
                        <label class="titulo">
                            <strong><span style="font-size: 6pt">PESO LÍQUIDO</span></strong>
                        </label>
                        <br />
                        <asp:TextBox ID="txtPesoLiquido" CssClass="txtDecimal3" runat="server" ForeColor="Blue"
                            BorderColor="White" Font-Names="Tahoma" />
                    </td>
                </tr>
                <tr>
                    <td colspan="10">
                        <table style="border: none; width: 100%;">
                            <tr>
                                <td valign="top" style="width: 33%; border: medium none;">
                                    <strong>DADOS ADICIONAIS</strong>
                                </td>
                                <td valign="top" style="width: 34%; border: medium none;">
                                    <strong>OBSERVAÇÕES FISCAIS</strong>
                                    <asp:Button ID="btnObservacoesFiscais" OnClick="btnObservacoesFiscais_Click" runat="server"
                                        UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />
                                </td>
                                <td valign="top" style="width: 33%; border: medium none;">
                                    <strong>OBSERVAÇÕES CONTROLE INTERNO</strong>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="10">
                        <table style="border: none; width: 100%;">
                            <tr>
                                <td valign="top" style="width: 33%; border: medium none;" align="center">
                                    <asp:TextBox ID="txtObservacoesDeEmbarque" runat="server" Width="98%" Height="155px"
                                        ForeColor="Blue" OnTextChanged="txtObservacoesDeEmbarque_TextChanged" Font-Names="Tahoma"
                                        TextMode="MultiLine" />
                                </td>
                                <td valign="top" style="width: 34%; border: medium none;" align="center">
                                    <asp:TextBox ID="txtObservacoesFiscais" runat="server" ClientIDMode="Static" Width="98%"
                                        Height="155px" ForeColor="Blue" OnTextChanged="txtObservacoesFiscais_TextChanged"
                                        Font-Names="Tahoma" TextMode="MultiLine" />
                                </td>
                                <td valign="top" style="width: 33%; border: medium none;" align="center">
                                    <asp:TextBox ID="txtObservacoesInternas" runat="server" Width="98%" Height="155px"
                                        ForeColor="Blue" OnTextChanged="txtObservacoesInternas_TextChanged" Font-Names="Tahoma"
                                        AutoPostBack="True" TextMode="MultiLine" />
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:consultaempresas id="ucConsultaEmpresas" runat="server" />
    <uc:consultaclientes id="ucConsultaClientes" runat="server" />
    <uc:pedidoxsaldo id="ucPedidoxSaldo" runat="server" />
    <uc:consultanotatroca id="ucConsultaNotaTroca" runat="server" />
    <uc:consultaromaneios id="ucConsultaRomaneios" runat="server" />
    <uc:consultaobservacoes id="ucConsultaObservacoes" runat="server" />
    <uc:consultaplacas id="ucConsultaPlacas" runat="server" />
    <uc:consultaestados id="ucConsultaEstados" runat="server" />
    <uc:consultacodmunicipios id="ucConsultaCodMunicipios" runat="server" />
    <uc:notafiscalxclassificacao id="ucNotaFiscalXClassificacao" runat="server" />
    <uc:consultanotavendaaordem id="ucConsultaNotaVendaAOrdem" runat="server" />
    <uc:consultaautorizacaoderetirada id="ucConsultaAutorizacaoDeRetirada" runat="server" />
    <uc:consultaencargosplanodecontas id="ucConsultaEncargosPlanoDeContas" runat="server" />
    <uc:consultadadosbancarios id="ucConsultaDadosBancarios" runat="server" />
    <uc:consultapedidosxnotas id="ucConsultaPedidosXNotas" runat="server" />
    <uc:notafiscalreferencial id="ucNotaFiscalReferencial" runat="server" />
    <uc:notadedevolucaoxnota id="ucNotaDeDevolucaoXNota" runat="server" />
    <uc:consultaprocuracao id="ucConsultaProcuracao" runat="server" />
    <uc:inutilizacao id="ucInutilizacao" runat="server" />
    <uc:vencimentos id="ucVencimentos" runat="server" />
    <uc:monitordenotas id="ucMonitorDeNotas" runat="server" />
    <uc:emailnfe id="ucEmailNFe" runat="server" />
    <uc:nfobsproduto id="ucNFObsProduto" runat="server" />
    <uc:nfencargo id="ucNFEncargo" runat="server" />
    <uc:consultaproduto id="ucConsultaProduto" runat="server" />
    <uc:nfprodutor id="ucNFProdutor" runat="server" />
    <uc:nfreferencialsaida id="ucNFReferencialSaida" runat="server" />
    <uc:consultadelote id="ucConsultaLote" runat="server" />
    <uc:consultarnaviosxinvoice id="ucConsultarNaviosXInvoice" runat="server" />
    <uc:consultartransferencias id="ucTransferencias" runat="server" />
    <uc:consultaliberacao id="ucLiberacao" runat="server" />

</asp:Content>
