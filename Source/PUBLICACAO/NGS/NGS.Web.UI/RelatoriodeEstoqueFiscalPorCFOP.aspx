<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatoriodeEstoqueFiscalPorCFOP.aspx.vb" Inherits="NGS.Web.UI.RelatoriodeEstoqueFiscalPorCFOP" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        //        function calcValorQtde() {
        //            $("#MainContent_ucPedidoXLancamento_updpnlPedidoLancamentoItem").mask("Carregando...");
        //            var valor = $("#MainContent_TabContainer1_TabSaldoInicial_txtEP").val();
        //            var total = $("#MainContent_TabContainer1_TabSaldoInicial_txtET").val();
        //            var parameters = "{ valor: '" + valor + "', total: '" + total + "', casasDecimais: '0' }";

        //            $.ajax({
        //                type: "POST",
        //                async: true,
        //                url: rootPath + "/WebMethods.asmx/calcTotal",
        //                data: parameters,
        //                contentType: "application/json; charset=utf-8",
        //                dataType: "json",
        //                success: function (msg) {
        //                    var valor = eval('(' + msg.d + ')');
        //                    $("#MainContent_TabContainer1_TabSaldoInicial_txtEG").val(valor);
        //                    $("#MainContent_updpnlform").unmask();
        //                },
        //                error: function (jqXHR, textStatus, errorThrown) {
        //                    msgbox("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')", "Possivel Erro Ocorrido!", "Erro");
        //                    $("#MainContent_updpnlform").unmask();
        //                }
        //            });
        //        }

        //        function calcValor() {
        //            $("#MainContent_ucPedidoXLancamento_updpnlPedidoLancamentoItem").mask("Carregando...");
        //            var valor = $("#MainContent_TabContainer1_TabSaldoInicial_txtEPValor").val();
        //            var total = $("#MainContent_TabContainer1_TabSaldoInicial_txtETValor").val();
        //            var parameters = "{ valor: '" + valor + "', total: '" + total + "', casasDecimais: '4' }";

        //            $.ajax({
        //                type: "POST",
        //                async: true,
        //                url: rootPath + "/WebMethods.asmx/calcTotal",
        //                data: parameters,
        //                contentType: "application/json; charset=utf-8",
        //                dataType: "json",
        //                success: function (msg) {
        //                    var valor = eval('(' + msg.d + ')');
        //                    $("#MainContent_TabContainer1_TabSaldoInicial_txtEGValor").val(valor);
        //                    $("#MainContent_updpnlform").unmask();
        //                },
        //                error: function (jqXHR, textStatus, errorThrown) {
        //                    msgbox("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')", "Possivel Erro Ocorrido!", "Erro");
        //                    $("#MainContent_updpnlform").unmask();
        //                }
        //            });
        //        }
    </script>
    <style type="text/css">
        .st1 {
            width: 113px;
        }

        .st2 {
            width: 165px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngForm" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlform" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Estoque Fiscal Por CFOP
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabRelatorio" runat="server" HeaderText="Relatório">
                    <ContentTemplate>
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
                                        </ul>
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjuda" Text="Ajudar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl st1">
                                <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." runat="server" Text="Empresa:" />
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" ToolTip="Empresa Para negociação/consulta." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl st1">
                                Parametro:
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="chkOcultarCFOP" runat="server" Text="Ocultar CFOP" data-ToolTip="default" ToolTip="Selecionar para ocultar o CFOP." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl st1" style="width: 113px;">
                                Ano:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlAno" runat="server" Width="75px" data-ToolTip="default" ToolTip="Ano de cadastro." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl st1">
                                Periodo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlMesInicial" runat="server" data-ToolTip="default" ToolTip="Mês inicial e dia final da consulta." />
                            </div>
                            <div class="coltxt">
                                à
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="106px" data-ToolTip="default" ToolTip="Mês inicial e dia final da consulta." />
                            </div>
                        </div>
                        <div class="row ">
                            <div class="coltxt lg">
                                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabSaldoInicial" runat="server" HeaderText="Saldo Inicial">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovo" runat="server" Text="Gravar" />
                                    </li>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultar" runat="server" Text="Consultar" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjuda2" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl st1">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresaSI" runat="server" Width="623px" data-ToolTip="default" ToolTip="Empresa Para negociação/consulta." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl st1">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeProduto" runat="server" Width="390px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="BtnProdutoSI" CssClass="btn" runat="server" Text=" &gt; " data-ToolTip="default" ToolTip="Nome do Produto." />
                            </div>
                            <div class="collbl st1">
                                Ano:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlAnoSI" runat="server" Width="65px" data-ToolTip="default" ToolTip="Ano da pesquisa." />
                            </div>
                        </div>
                        <div class="painelleft" style="width: 34%;">
                            <div class="subtitulodiv">
                                Quantidades
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Em Nosso Poder(NP):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNP" CssClass="txtInteiro" runat="server" Width="120px" Style="text-align: right;" data-ToolTip="default" ToolTip="Quantidade em poder da empresa." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Em Poder de Terceiros (PT):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPT" CssClass="txtInteiro" runat="server" Style="width: 120px; text-align: right;" data-ToolTip="default" ToolTip="Quantidade em poder de outras empresas." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Para Fins de Exportação (FE):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtFE" CssClass="txtInteiro" runat="server" Style="width: 120px; text-align: right;" data-ToolTip="default" ToolTip="Quantidade a ser esportada." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Estoque Proprio(EP):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtEP" CssClass="txtInteiro" runat="server" Style="width: 120px; text-align: right;" data-ToolTip="default" ToolTip="Quantidade de estoque da própria empresa." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Estoque de Terceiros(ET):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtET" CssClass="txtInteiro" runat="server" Style="width: 120px; text-align: right;" data-ToolTip="default" ToolTip="Quantidade de estoque de outras empresas." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Estoque Geral(EG):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtEG" CssClass="txtInteiro" runat="server" Style="width: 120px; text-align: right;" data-ToolTip="default" ToolTip="Quantidade do estoque geral." />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 34%; margin-left: 2px;">
                            <div class="subtitulodiv">
                                Valores
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Em Nosso Poder(NP):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNPValor" CssClass="txtDecimal4" runat="server" Width="120px"
                                        Style="text-align: right;" data-ToolTip="default" ToolTip="Valor em poder da empresa." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Em Poder de Terceiros (PT):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPTValor" CssClass="txtDecimal4" runat="server" Style="width: 120px; text-align: right;" data-ToolTip="default" ToolTip="Valor em poder de outras empresas." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Para Fins de Exportação (FE):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtFEValor" CssClass="txtDecimal4" runat="server" Style="width: 120px; text-align: right;" data-ToolTip="default" ToolTip="Valor a ser esportada." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Estoque Proprio(EP):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtEPValor" CssClass="txtDecimal4" runat="server" Style="width: 120px; text-align: right;" data-ToolTip="default" ToolTip="Valor de estoque da própria empresa." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Estoque de Terceiros(ET):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtETValor" CssClass="txtDecimal4" runat="server" Style="width: 120px; text-align: right;" data-ToolTip="default" ToolTip="Valor de estoque de outras empresas." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl st2">
                                    Estoque Geral(EG):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtEGValor" CssClass="txtDecimal4" runat="server" Style="width: 120px; text-align: right;" data-ToolTip="default" ToolTip="Valor do estoque geral." />
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="subtitulodiv">
                                Saldos Iniciais Gravados
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 300px;">
                            <asp:GridView ID="gridSI" runat="server" CellPadding="4" ForeColor="#333333" GridLines="None"
                                AutoGenerateColumns="False" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" Width="100%" />
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
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                                    <asp:BoundField DataField="Empresa" HeaderText="Empresa" ReadOnly="True">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Ano" HeaderText="Ano">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Produto" HeaderText="Produto" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EmNossoPoder" HeaderText="NP" HtmlEncode="False" DataFormatString="{0:N0}">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EmPoderDeTerceiros" HeaderText="PT" DataFormatString="{0:N0}">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ParaFinsDeExportacao" HeaderText="FE" DataFormatString="{0:N0}">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="TotalEstoqueProprio" HeaderText="EP" DataFormatString="{0:N0}">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EstoqueDeTerceiro" HeaderText="ET" DataFormatString="{0:N0}">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="TotalEstoqueGeral" HeaderText="EG" DataFormatString="{0:N0}">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>
