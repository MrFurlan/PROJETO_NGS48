<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="MDFe.aspx.vb" Inherits="NGS.Web.UI.MDFe" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }

        input[type='text'] {
            height: 16px;
        }

        select {
            height: 20px;
        }

        .collbl {
            width: 150px;
        }

        fieldset {
            border-radius: 5px !important;
            padding: 5px !important;
            border: 1px solid #BBBBBB;
            margin-top: 4px !important;
        }

            fieldset legend {
                font-family: Tahoma,Arial,Helvetica,sans-serif;
                font-size: 0.8em;
                font-weight: bold;
                padding-left: 5px;
                padding-right: 5px;
            }

            fieldset td {
                vertical-align: top;
                white-space: nowrap;
            }

            fieldset label {
                font-size: 7pt;
                font-weight: bold;
            }

            fieldset .coltxt span, fieldset .coltxt input[type=text]:not(.no) {
                font-size: 7pt;
                color: blue;
            }

        .coltxt {
            min-height: 28px;
        }

        .aspNetDisabled {
            cursor: no-drop;
        }
    </style>
    <script type="text/javascript">
        function pageLoadManifestoEletronicoDeDocumentosFiscais() {
            $("input[type=checkbox]:visible", "#MainContent_TabContainer1_tabNotaFiscal_grdNFeSaldo").change(function () {
                if ($(this).is(":checked") == true || $(this).is(":checked") == "true") {
                    $(this).parent().parent().parent().find("input[type=text].txtInteiro").prop('disabled', false).setMask("integer");
                } else {
                    $(this).parent().parent().parent().find("input[type=text].txtInteiro").prop('disabled', true).setMask("integer");
                }
            });
        }

        $(document).ready(function () {
            pageLoadManifestoEletronicoDeDocumentosFiscais();
            var prmManifestoEletronicoDeDocumentosFiscais = Sys.WebForms.PageRequestManager.getInstance();
            prmManifestoEletronicoDeDocumentosFiscais.add_endRequest(pageLoadManifestoEletronicoDeDocumentosFiscais);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmManifestoEletronicoDeDocumentosFiscais" runat="server"
        AsyncPostBackTimeout="1000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlManifestoEletronicoDeDocumentosFiscais" runat="server">
        <ContentTemplate>
            <script type="text/javascript">
                function selectAllCheckboxes(chkAll) {
                    var chk = $('#' + chkAll.id);
                    var checked = chk.attr('checked') == "checked";
                    $("input[type='checkbox']", "#MainContent_TabContainer1_tabNotaFiscal_grdNFe").not("#chkAll").each(function () {
                        $(this).attr("checked", checked);
                    });
                }
            </script>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdfCodigo" runat="server" />
            <div class="titulodiv">
                <label>Manifesto Eletrônico de Documentos Fiscais (MDF-e)</label>
                <div style="float: right; padding-right: 10px;">
                    <asp:ImageButton ID="imgUsuario" runat="server" Width="18px" Height="20px" ImageUrl="~/images/man2.png"
                        ImageAlign="AbsMiddle" Style="border: 0;"></asp:ImageButton>
                    <asp:Label ID="lblUsuario" runat="server" Font-Bold="True" Font-Size="11px" Style="font-family: Tahoma,Arial,Helvetica,sans-serif;" />
                </div>
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" Width="100%" ActiveTabIndex="1">
                <ajaxToolkit:TabPanel ID="tabNotaFiscal" runat="server" text="Nota Fiscal">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultarC" runat="server" Text="Consultar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkSelecionarC" runat="server" Text="Selecionar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparC" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaC" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresa" runat="server" Width="600px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                                <asp:TextBox ID="txtNomeCliente" runat="server" Width="560px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnCliente" runat="server" Text=">" OnClick="btnCliente_Click" CssClass="btn"
                                    data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Número NF:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNumMDFe" runat="server" Width="230px" data-ToolTip="default"
                                    ToolTip="Inserir o número da Nota Fiscal." />
                            </div>
                            <div class="collbl">
                                Notas de:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdoEntrada" runat="server" Text="Entradas" GroupName="Radio"
                                    data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdoSaida" runat="server" Text="Saídas" Checked="True" GroupName="Radio"
                                    data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data Inicial:
                            </div>
                            <div class="coltxt" style="width: 240px;">
                                <asp:TextBox ID="txtPeriodoInicial" runat="server" Width="100px" CssClass="calendario"
                                    data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                            </div>
                            <div class="collbl">
                                Data Final:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPeriodoFinal" runat="server" Width="100px" CssClass="calendario"
                                    data-ToolTip="default" ToolTip="Período Final da pesquisa." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="grdNFe" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" ShowHeaderWhenEmpty="True"
                                EmptyDataText="NENHUM REGISTRO ENCONTRADO!">
                                <EditRowStyle BackColor="#999999" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EmptyDataRowStyle VerticalAlign="Middle" HorizontalAlign="Center" Font-Bold="True" />
                                <Columns>
                                    <asp:TemplateField>
                                        <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <HeaderTemplate>
                                            <input id="chkAll" onclick="selectAllCheckboxes(this);" type="checkbox" />
                                        </HeaderTemplate>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkMDFe" runat="server" CssClass="ckCkeck" Checked='<%# Bind("chkMDFe") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="EntradaSaida_Id" HeaderText="E/S">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nota_Id" HeaderText="Nota">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Serie_Id" HeaderText="Série">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente_Id" HeaderText="Cliente">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EndCliente_Id" HeaderText="End">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nome" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Operacao" HeaderText="OP">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="SubOperacao" HeaderText="SO">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Data" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="PesoFiscal" DataFormatString="{0:N0}" HeaderText="Peso">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="tabMDFe" runat="server">
                    <HeaderTemplate>
                        Emissão MDF-e
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server" class="iconNovo">
                                        <asp:LinkButton Text="Gravar" ID="lnkNovo" runat="server" />
                                    </li>
                                    <li runat="server" class="iconExcluir">
                                        <asp:LinkButton ID="lnkExcluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" Text="Excluir" />
                                    </li>
                                    <li runat="server" class="iconConsultar">
                                        <asp:LinkButton ID="lnkConsultar" runat="server" Text="Consultar" />
                                    </li>
                                    <li runat="server" class="iconLimpar">
                                        <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server" class="iconRelatorio">
                                        <asp:LinkButton ID="lnkEspelho" runat="server" Text="Espelho" />
                                    </li>
                                    <li runat="server" class="iconRelatorio">
                                        <asp:LinkButton ID="lnkDAMDFE" runat="server" Text="DAMDFE" />
                                    </li>
                                    <li runat="server" class="iconConfirmar">
                                        <asp:LinkButton ID="lnkEncerrar" runat="server" Text="Encerrar" />
                                    </li>
                                    <li runat="server" class="iconRelatorio">
                                        <asp:LinkButton ID="lnkEnviarSEFAZ" runat="server" Text="Enviar SEFAZ" />
                                    </li>
                                    <li runat="server" class="iconMail">
                                        <asp:LinkButton ID="lnkEnviarEmail" runat="server" Text="Enviar E-mail" />
                                    </li>
                                    <li class="iconConsultar" runat="server">
                                        <asp:UpdatePanel ID="updpnlVisualizar" runat="server">
                                            <ContentTemplate>
                                                <asp:LinkButton ID="lnkVisualizar" Text="Pré-Visualizar" runat="server" />
                                            </ContentTemplate>
                                            <Triggers>
                                                <asp:PostBackTrigger ControlID="lnkVisualizar" />
                                            </Triggers>
                                        </asp:UpdatePanel>
                                    </li>
                                    <li runat="server" class="iconAjuda">
                                        <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 70%;">
                            <fieldset>
                                <legend>EMPRESA</legend>
                                <table cellspacing="0" cellpadding="0" width="100%">
                                    <tr>
                                        <td>
                                            <label>NOME/RAZÃO SOCIAL</label><br />
                                            <div class="coltxt">
                                                <asp:Button ID="btnEmpresa" runat="server" Text=">" UseSubmitBehavior="False" CssClass="btn" />
                                                <asp:Label ID="txtNomeDaEmpresa" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>ENDEREÇO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEnderecoDaEmpresa" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>COMPLEMENTO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtComplementoDaEmpresa" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>BAIRRO/DISTRITO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtBairroDaEmpresa" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>CEP</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCepDaEmpresa" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>INSCRIÇÃO ESTADUAL</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtInscricaoDaEmpresa" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>CNPJ</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCnpjDaEmpresa" runat="server" />
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </fieldset>
                            <fieldset>
                                <legend>REMETENTE</legend>
                                <table cellspacing="0" cellpadding="0" width="100%">
                                    <tr>
                                        <td>
                                            <label>NOME/RAZÃO SOCIAL</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtNomeDoRemetente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>ENDEREÇO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEnderecoDoRemetente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>COMPLEMENTO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtComplementoDoRemetente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>BAIRRO/DISTRITO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtBairroDoRemetente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>CEP</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCepDoRemetente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>INSCRIÇÃO ESTADUAL</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtInscricaoDoRemetente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>CNPJ</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCnpjDoRemetente" runat="server" />
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </fieldset>
                            <fieldset>
                                <legend>DESTINATÁRIO</legend>
                                <table cellspacing="0" cellpadding="0" width="100%">
                                    <tr>
                                        <td>
                                            <label>NOME/RAZÃO SOCIAL</label><br />
                                            <div class="coltxt">
                                                <asp:Button ID="btnClienteMDFe" runat="server" Text=">" UseSubmitBehavior="False" CssClass="btn"
                                                    OnClick="btnClienteMDFe_Click" />
                                                <asp:Label ID="txtNomeDoCliente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>CNPJ/CPF</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCnpjDoCliente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>INSCRIÇÃO ESTADUAL</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtInscricaoDoCliente" runat="server" />

                                            </div>
                                        </td>
                                        <td style="width: 115px;">
                                            <label>DATA DA EMISSÃO</label>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtDataDeEmissao" runat="server" Width="60px" BorderStyle="None" CssClass="calendario" />
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                                <table cellspacing="0" cellpadding="0" width="100%">
                                    <tr>
                                        <td>
                                            <label>ENDEREÇO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEnderecoDoCliente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>COMPLEMENTO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtComplementoDoCliente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>BAIRRO/DISTRITO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtBairroDoCliente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>CEP</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCepDoCliente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>MUNICÍPIO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCidadeDoCliente" runat="server" />
                                            </div>

                                        </td>
                                        <td>
                                            <label>UF</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEstadoDoCliente" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>FONE/FAX</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtTelefoneDoCliente" runat="server" />
                                            </div>
                                        </td>
                                        <td style="width: 115px;">
                                            <label>DATA DE SAÍDA</label>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtDataDeSaida" runat="server" Width="60px" BorderStyle="None" CssClass="calendario" />
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </fieldset>
                        </div>
                        <div class="painelright" style="width: 30%;">
                            <fieldset>
                                <legend>DADOS DO MDF-e</legend>
                                <table cellspacing="0" cellpadding="0" style="width: 100%;">
                                    <tr>
                                        <td colspan="2" valign="top" align="center">
                                            <label>
                                                <strong><span style="font-size: 8px;">CHAVE DE ACESSO DO MDF-e CONSULTA NO SITE: </span><a target="_trackPageview" href="http://WWW.NFE.FAZENDA.GOV.BR">WWW.NFE.FAZENDA.GOV.BR</a></strong>
                                            </label>
                                            <div class="coltxt" style="width: 99%;">
                                                <asp:TextBox ID="txtChaveNFe" CssClass="no" runat="server" Width="99%" Enabled="False" MaxLength="50" Style="margin: 3px 0;" />
                                                <ajaxToolkit:MaskedEditExtender ID="mkeChaveNFe" runat="server" TargetControlID="txtChaveNFe"
                                                    MaskType="Number" Mask="9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999"
                                                    Century="2000" CultureAMPMPlaceholder="" CultureCurrencySymbolPlaceholder=""
                                                    CultureDateFormat="" CultureDatePlaceholder="" CultureDecimalPlaceholder="" CultureThousandsPlaceholder=""
                                                    CultureTimePlaceholder="" BehaviorID="_content_mkeChaveNFe"></ajaxToolkit:MaskedEditExtender>
                                            </div>
                                        </td>
                                    </tr>
                                    <tr style="padding-top: 4px; border-top: solid 2px;">
                                        <td align="left" style="width: 115px; padding-left: 9px;">
                                            <label>
                                                <strong><span style="font-size: 10pt;">NÚMERO:
                                            </label>
                                        </td>
                                        <td align="right">
                                            <asp:TextBox ID="txtNumero" runat="server" Font-Size="12pt" Height="24px"
                                                Width="220px" Style="text-align: right;" MaxLength="10" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="left" style="padding-left: 9px;">
                                            <label>
                                                <strong><span style="font-size: 10pt;">SÉRIE:
                                            </label>
                                        </td>
                                        <td align="right">
                                            <asp:TextBox ID="txtSerie" runat="server" Font-Size="12pt" Height="24px"
                                                Width="220px" Style="text-align: right;" MaxLength="3" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="left" style="padding-left: 9px;">
                                            <label>
                                                <strong><span style="font-size: 10pt;">PESO BRUTO:
                                            </label>
                                        </td>
                                        <td align="right">
                                            <asp:TextBox ID="txtPesoBruto" runat="server" CssClass="txtDecimal"
                                                Font-Size="12pt" Height="24px" Width="220px"
                                                Style="text-align: right;" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="left" style="padding-left: 9px;">
                                            <label>
                                                <strong><span style="font-size: 10pt;">TOTAL:
                                            </label>
                                        </td>
                                        <td align="right">
                                            <asp:TextBox ID="txtValor" runat="server" CssClass="txtDecimal"
                                                Font-Size="12pt" Height="24px" Width="220px"
                                                Style="text-align: right;" Enabled="False" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="left" style="padding-left: 9px;">
                                            <label>
                                                <strong><span style="font-size: 10pt; color: black;">OPERAÇÃO:
                                            </label>
                                        </td>
                                        <td align="right" style="vertical-align: middle;">
                                            <asp:TextBox ID="txtOperacao" runat="server" Font-Size="9pt"
                                                Height="24px" Width="220px" Style="vertical-align: middle;" Enabled="False" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="left" style="padding-left: 9px;">
                                            <label>
                                                <strong><span style="font-size: 10pt; color: black;">CFOP:</label>
                                        </td>
                                        <td align="right">
                                            <asp:TextBox ID="txtCfop" runat="server" Font-Size="7pt"
                                                Height="24px" Width="220px" Style="vertical-align: middle;" Enabled="False" />
                                        </td>
                                    </tr>
                                </table>
                            </fieldset>
                        </div>
                        <div class="painelleft" style="width: 100%;">
                            <fieldset>
                                <legend>NOTAS FISCAIS</legend>
                                <div class="bordagrid" style="height: 135px;">
                                    <asp:GridView ID="gridNotasFiscais" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                        ForeColor="#333333" GridLines="None" Width="100%">
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <EditRowStyle BackColor="#999999" />
                                        <Columns>
                                            <asp:BoundField DataField="Nota" HeaderText="Nota">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="CodigoOperacao" HeaderText="Opera&#231;&#227;o">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Produto" HeaderText="Produto">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="CFOP" HeaderText="CFOP">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="QuantidadeFisica" HeaderText="F&#237;sico" DataFormatString="{0:N0}">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="QuantidadeFiscal" HeaderText="Fiscal" DataFormatString="{0:N0}">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ValorUnitario" HeaderText="Unit&#225;rio" DataFormatString="{0:N10}">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ValorTotal" HeaderText="Total" DataFormatString="{0:N2}">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </fieldset>
                        </div>
                        <div class="painelright" style="width: 30%;">
                            <fieldset>
                                <legend>OBSERVAÇÃO</legend>
                                <div class="subtitulodiv">
                                    OBSERVAÇÃO
                                </div>
                                <div class="row">
                                    <div class="coltxt" style="width: 97%;">
                                        <asp:TextBox ID="txtObservacao" runat="server" Height="215px" TextMode="MultiLine"
                                            Width="100%" />
                                    </div>
                                </div>
                            </fieldset>
                        </div>
                        <div class="painelleft" style="width: 25%;">
                            <fieldset style="height: 91px;">
                                <legend>PERCURSOS</legend>
                                <label>ESTADOS</label><br />
                                <div class="coltxt">
                                    <asp:Button ID="btnEstados" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False" OnClick="btnEstados_Click" />
                                    <asp:Label ID="lblEstados" runat="server" />
                                </div>
                            </fieldset>
                        </div>
                        <div class="painelleft" style="width: 45%;">
                            <fieldset>
                                <legend>LOCAL DE EMBARQUE/COLETA</legend>
                                <table cellspacing="0" cellpadding="0" width="100%">
                                    <tr>
                                        <td>
                                            <label>LOCAL DE EMBARQUE/COLETA</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEmbarque" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>MUNICÍPIO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtMunicipioEmbarque" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>UF</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEstadoEmbarque" runat="server" />
                                            </div>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <label>LOCAL DE DESCARGA/DESTINO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtNomeOrigemDestino" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>MUNICÍPIO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtMunicipioOrigemDestino" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>UF</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEstadoOrigemDestino" runat="server" />
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </fieldset>
                        </div>
                        <div class="painelleft" style="width: 70%;">
                            <fieldset>
                                <legend>DADOS DO TRANSPORTADOR</legend>
                                <table cellspacing="0" cellpadding="0" width="100%">
                                    <tr>
                                        <td>
                                            <label>RAZÃO SOCIAL</label><br />
                                            <div class="coltxt">
                                                <asp:Button ID="btnTransportador" runat="server" Text=">" UseSubmitBehavior="False" CssClass="btn"
                                                    OnClick="btnTransportador_Click" />
                                                <asp:Label ID="txtNomeDoTransportador" runat="server" Width="260px" />

                                            </div>
                                        </td>
                                        <td>
                                            <label>CÓDIGO ANTT</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtRNTRCTransportador" runat="server" />
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                                <table cellspacing="0" cellpadding="0" width="100%">
                                    <tr>
                                        <td>
                                            <label>ENDEREÇO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEnderecoDoTransportador" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>MUNICÍPIO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCidadeDoTransportador" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>UF</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEstadoDoTransportador" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>CNPJ/CPF</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCnpjDoTransportador" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>INSCRIÇÃO ESTADUAL</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtInscricaoDoTransportador" runat="server" />
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </fieldset>
                            <fieldset>
                                <legend>DADOS DO MOTORISTA</legend>
                                <table cellspacing="0" cellpadding="0" width="100%">
                                    <tr>
                                        <td>
                                            <label>NOME</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtNomeDoMotorista" runat="server" Width="260px" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>ENDEREÇO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEnderecoDoMotorista" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>MUNICÍPIO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCidadeDoMotorista" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>UF</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEstadoDoMotorista" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>CPF</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCPFDoMotorista" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>HABILITAÇÃO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtHabilitacaoDoMotorista" runat="server" />
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </fieldset>
                            <fieldset>
                                <legend>DADOS DO CAMINHÃO</legend>
                                <table cellspacing="0" cellpadding="0" width="100%">
                                    <tr>
                                        <td style="width: 100px;">
                                            <label>PLACA 1</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCPlaca1" runat="server" />
                                            </div>
                                        </td>
                                        <td style="width: 125px;">
                                            <label>MUNICÍPIO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCidadePlaca1" runat="server" />
                                            </div>
                                        </td>
                                        <td style="width: 85px;">
                                            <label>UF</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEstadoPlaca1" runat="server" />
                                            </div>
                                        </td>
                                        <td style="width: 85px;">
                                            <label>RNTRC</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtRNTRCPlaca1" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>PROPRIETÁRIO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtProprietarioPlaca1" runat="server" />
                                            </div>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <label>PLACA 2</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCPlaca2" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>MUNICÍPIO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCidadePlaca2" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>UF</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEstadoPlaca2" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>RNTRC</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtRNTRCPlaca2" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>PROPRIETÁRIO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtProprietarioPlaca2" runat="server" />
                                            </div>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <label>PLACA 3</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCPlaca3" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>MUNICÍPIO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtCidadePlaca3" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>UF</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtEstadoPlaca3" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>RNTRC</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtRNTRCPlaca3" runat="server" />
                                            </div>
                                        </td>
                                        <td>
                                            <label>PROPRIETÁRIO</label>
                                            <div class="coltxt">
                                                <asp:Label ID="txtProprietarioPlaca3" runat="server" />
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </fieldset>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaClientesDireto ID="ucConsultaClientesDireto" runat="server" />
    <uc:ConsultaMDFeXNotas ID="ucConsultaMDFeXNotas" runat="server" />
    <uc:ConsultaPlacas ID="ucConsultaPlacas" runat="server" />
    <uc:ConsultaEstados ID="ucConsultaEstados" runat="server" />
    <uc:MDFeXEstado ID="ucMDFeXEstado" runat="server" />
    <uc:EmailNFe ID="ucEmailNFe" runat="server" />
</asp:Content>
