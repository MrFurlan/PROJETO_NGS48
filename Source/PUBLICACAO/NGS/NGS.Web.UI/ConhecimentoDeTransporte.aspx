<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ConhecimentoDeTransporte.aspx.vb" Inherits="NGS.Web.UI.ConhecimentoDeTransporte" %>

<%@ Import Namespace="NGS.Lib.Negocio" %>
<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        #meioconteudo {
            width: 1320px !important;
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
    </style>
    <script type="text/javascript">
        function pageLoadConhecimentoDeTransporte() {
            $("input[type=checkbox]:visible", "#MainContent_tcPrincipal_tabNotaFiscal_grdNFeSaldo").change(function () {
                if ($(this).is(":checked") == true || $(this).is(":checked") == "true") {
                    $(this).parent().parent().parent().find("input[type=text].txtInteiro").prop('disabled', false).setMask("integer");
                } else {
                    $(this).parent().parent().parent().find("input[type=text].txtInteiro").prop('disabled', true).setMask("integer");
                }
            });


            $("#MainContent_tcPrincipal_tabCTRC_txtChaveNFe").live('keypress', function (e) {
                var code = (e.keyCode ? e.keyCode : e.which);
                if (code == 13) {
                    $("#MainContent_tcPrincipal_tabCTRC_imgVerificarChaveNFE").click();
                    return false;
                }
            });
        }

        $(document).ready(function () {
            pageLoadConhecimentoDeTransporte();
            var prmConhecimentoDeTransporte = Sys.WebForms.PageRequestManager.getInstance();
            prmConhecimentoDeTransporte.add_endRequest(pageLoadConhecimentoDeTransporte);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngConhecimentoDeTransporte" runat="server" AsyncPostBackTimeout="1000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlConhecimentoDeTransporte" runat="server">
        <ContentTemplate>
            <script type="text/javascript">
                function selectAllCheckboxes1(chkAll) {
                    var chk = $('#' + chkAll.id);
                    var checked = chk.attr('checked') == "checked";
                    $("input[type='checkbox']", "#MainContent_tcPrincipal_tabNotaFiscal_grdNFe").not("#chkAll").each(function () {
                        $(this).attr("checked", checked);
                    });
                }

                function selectAllCheckboxes2(chkAll) {
                    var chk = $('#' + chkAll.id);
                    var checked = chk.attr('checked') == "checked";
                    $("input[type='checkbox']", "#MainContent_tcPrincipal_tabNotaFiscal_grdNFeSaldo").not("#chkAll").each(function () {
                        $(this).attr("checked", checked);
                    });
                }

                function somaSelecionados(ischk) {
                    var chk = $('#' + ischk.id);
                    var checked = chk.attr('checked') == "checked";

                    var txtValorFrete = $(ischk).parent().parent().parent().parent().find(".txtValorFrete");
                    var txtPeso = $(ischk).parent().parent().parent().parent().find(".txtPeso");

                    var txtPesoTotal = $("#MainContent_tcPrincipal_tabNotaFiscal_txtPesoTotalSelecionado");
                    var txtValorTotalFrete = $("#MainContent_tcPrincipal_tabNotaFiscal_txtValorFreteSelecionado");

                    //RPA
                    if ((txtPeso.text() != undefined) && (txtValorFrete.val() != undefined)) {

                        if ((txtPesoTotal.val() != undefined) && (txtValorTotalFrete.val() != undefined)) {

                            var pesoTotal = parseInt(txtPesoTotal.val().split(".").join(""));
                            var valorTotal = parseFloat(txtValorTotalFrete.val().split(".").join("").split(",").join("."));

                            if (checked) {
                                pesoTotal += parseInt(txtPeso.val().split(".").join(""));
                                valorTotal += parseFloat(txtValorFrete.val().split(".").join(""));
                            } else {
                                pesoTotal -= parseInt(txtPeso.val().split(".").join(""));
                                valorTotal -= parseFloat(txtValorFrete.val().split(".").join(""));
                            }
                            txtPesoTotal.val(float2Moeda(pesoTotal, false));
                            txtValorTotalFrete.val(float2Moeda(valorTotal.toFixed(2), true));
                        }
                    }
                }

            </script>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdfTipo" runat="server" />
            <asp:HiddenField ID="hdfCodigo" runat="server" />
            <asp:HiddenField ID="hdfViaDeTransporte" runat="server" />
            <table style="width: 100%; padding: 0px; border: 0px;">
                <tr>
                    <td>
                        <div class="titulodiv">
                            <div style="float: left;">
                                Conhecimento de Transporte
                            </div>
                            <div style="float: right; padding-right: 10px;">
                                <asp:ImageButton ID="imgUsuario" runat="server" Width="18px" Height="20px" ImageUrl="~/images/man2.png"
                                    ImageAlign="AbsMiddle" Style="border: 0;"></asp:ImageButton>
                                <asp:Label ID="lblUsuario" runat="server" Font-Bold="True" Font-Size="11px" Style="font-family: Tahoma,Arial,Helvetica,sans-serif;" />
                            </div>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>
                        <ajaxToolkit:TabContainer ID="tcPrincipal" runat="server" ActiveTabIndex="0" Width="100%">
                            <ajaxToolkit:TabPanel ID="tabNotaFiscal" runat="server">
                                <HeaderTemplate>
                                    Emissão CTe
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="menu_acoes">
                                        <div class="acoes">
                                            <ul>
                                                <li runat="server">
                                                    <asp:LinkButton ID="lnkConsultarNota" Text="Consultar" runat="server" class="iconConsultar" />
                                                </li>
                                                <li runat="server">
                                                    <asp:LinkButton ID="lnkSelecionarNota" Text="Selecionar" runat="server" CssClass="iconNovo" />
                                                </li>
                                                <li runat="server">
                                                    <asp:LinkButton ID="lnkRelatorioNota" Text="Relatório" runat="server" CssClass="iconRelatorio" />
                                                </li>
                                                <li runat="server">
                                                    <asp:LinkButton ID="lnkLimparNota" Text="Limpar" runat="server" CssClass="iconLimpar" />
                                                </li>
                                                <li runat="server">
                                                    <asp:LinkButton ID="lnkAjudaNota" Text="Ajuda" runat="server" CssClass="iconAjuda" />
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Empresa:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="DdlEmpresa" runat="server" Width="600px" Style="font-family: Tahoma,Arial,Helvetica,sans-serif; font-size: 11px;" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Cliente:
                                        </div>
                                        <div class="coltxt">
                                            <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                                            <asp:TextBox ID="txtNomeCliente" runat="server" Width="570px" Enabled="False" />
                                            <asp:Button ID="btnCliente" runat="server" Text=">" CssClass="btn" OnClick="btnCliente_Click"
                                                data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Pedido:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtPedido" runat="server" Width="90px" CssClass="txtNumerico" Style="text-align: right"
                                                data-ToolTip="default" ToolTip="Informar o número do pedido." />
                                        </div>
                                        <div class="coltxt">
                                            <asp:Button ID="btnPedido" runat="server" Text=">" UseSubmitBehavior="False" CssClass="btn"
                                                data-ToolTip="default" ToolTip="Informar o número do pedido." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Número NF:
                                            <asp:CheckBox ID="chkNFxCtes" runat="server" BorderStyle="None" Checked="False" data-ToolTip="default"
                                                ToolTip="Esta nota será utilizada para mais de um CTE" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtNumConhecimento" runat="server" Width="100px" Style="text-align: right"
                                                data-ToolTip="default" ToolTip="Inserir o número da Nota Fiscal." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Notas de:
                                        </div>
                                        <div class="coltxt">
                                            <asp:RadioButton ID="rdoEntrada" runat="server" Text="Entradas" GroupName="Radio"
                                                data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                                            <asp:RadioButton ID="rdoSaida" runat="server" Text="Saídas" Checked="True" GroupName="Radio"
                                                data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Via de Transporte:
                                        </div>
                                        <div class="coltxt">
                                            <asp:RadioButton ID="rdoRodoviario" runat="server" Text="Rodoviário" GroupName="ViaDeTransporte"
                                                Checked="True" data-ToolTip="default" ToolTip="Escolher o meio de transporte, rodoviário ou ferroviário." />
                                            <asp:RadioButton ID="rdoFerroviario" runat="server" Text="Ferroviário" GroupName="ViaDeTransporte"
                                                data-ToolTip="default" ToolTip="Escolher o meio de transporte, rodoviário ou ferroviário." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Tipo de Conhecimento:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlTipoConhecimento" runat="server" Width="600px" Style="font-family: Tahoma,Arial,Helvetica,sans-serif; font-size: 11px;"
                                                AutoPostBack="True" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Data Inicial:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtPeriodoInicial" runat="server" Width="100px" CssClass="calendario"
                                                data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                                        </div>
                                        <div class="collbl">
                                            Data Final:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtPeriodoFinal" runat="server" Width="100px" CssClass="calendario"
                                                data-ToolTip="default" ToolTip="Período final da pesquisa." />
                                        </div>
                                    </div>
                                    <div class="row" runat="server" id="divClienteRPA">
                                        <div class="collbl">
                                            Cliente RPA:
                                        </div>
                                        <div class="coltxt">
                                            <asp:HiddenField ID="txtCodigoClienteRPA" runat="server" />
                                            <asp:TextBox ID="txtNomeClienteRPA" runat="server" Width="570px" Enabled="False" />
                                            <asp:Button ID="btnClienteRPA" runat="server" Text=">" />
                                        </div>
                                    </div>
                                    <div class="row" runat="server" id="divPaineltotais" style="display: none;">
                                        <div class="collbl" style="width: 150px">
                                            Peso Selecionado:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox runat="server" ID="txtPesoTotalSelecionado" CssClass="txtInteiro" Enabled="False"
                                                Text="0" Font-Bold="True" />
                                        </div>
                                        <div class="collbl" style="width: 200px">
                                            Valor do Frete Selecionado:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox runat="server" ID="txtValorFreteSelecionado" CssClass="txtDecimal" Enabled="False"
                                                Text="0,00" Font-Bold="True" />
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
                                                        <input id="chkAll" onclick="selectAllCheckboxes1(this);" type="checkbox" />
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:CheckBox ID="chkConhecimento" runat="server" CssClass="ckCkeck" Checked='<%# Bind("chkConhecimento") %>'
                                                            OnCheckedChanged="chkConhecimento_CheckedChanged" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="EntradaSaida_Id" HeaderText="E/S">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Nota_Id" HeaderText="Nota">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
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
                                                <asp:TemplateField HeaderText="Peso NF">
                                                    <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtPesoFiscal" runat="server" Width="80px" CssClass="txtInteiro txtPeso"
                                                            Enabled="false" Text='<%# String.Format("{0:N0}", Eval("PesoFiscal"))%>' />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:TemplateField HeaderText="Peso de Chegada">
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtPesoDeChegada" runat="server" Width="80px" CssClass="txtInteiro txtPeso"
                                                            Enabled="false" Text='<%# String.Format("{0:N0}", Eval("PesoDeChegada"))%>' />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="TarifaFrete" DataFormatString="{0:N2}" HeaderText="Tarifa Frete">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:TemplateField HeaderText="Valor Frete">
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtValorFrete" runat="server" Width="80px" CssClass="txtDecimal txtValorFrete"
                                                            Enabled="false" Text='<%# String.Format("{0:N2}", Eval("ValorFrete"))%>' />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="Saldo" DataFormatString="{0:N0}" HeaderText="Saldo Qtd.">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                            </Columns>
                                        </asp:GridView>
                                        <asp:GridView ID="grdNFeSaldo" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                            ForeColor="#333333" GridLines="None" Width="100%" ShowHeaderWhenEmpty="True"
                                            EmptyDataText="NENHUM REGISTRO ENCONTRADO!" Visible="False">
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
                                                        <input id="chkAll" onclick="selectAllCheckboxes2(this);" type="checkbox" />
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:CheckBox ID="chkConhecimento" runat="server" CssClass="ckCkeck" Checked='<%# Bind("chkConhecimento") %>' />
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
                                                    <HeaderStyle HorizontalAlign="Left" Width="275px" Wrap="True"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left" Width="275px" Wrap="True"></ItemStyle>
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
                                                <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="PesoFiscal" DataFormatString="{0:N0}" HeaderText="Peso">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="ConsumidoPeso" DataFormatString="{0:N0}" HeaderText="Consumido">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="PesoSaldo" DataFormatString="{0:N0}" HeaderText="Saldo">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:TemplateField HeaderText="Peso Kg">
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtVlrPeso" Style="text-align: right;" CssClass="txtInteiro" Text='<%# String.Format("{0:N0}", Eval("PesoSaldo"))%>'
                                                            runat="server" BorderColor="White" Width="100px" Enabled="false" />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:TemplateField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="tabCTRC" runat="server">
                                <HeaderTemplate>
                                    Conhecimento de Transporte
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="menu_acoes">
                                        <div class="acoes">
                                            <ul>
                                                <li runat="server" class="iconNovo">
                                                    <asp:LinkButton ID="lnkEmitir" Text="Emitir" runat="server" />
                                                </li>
                                                <li runat="server" class="iconNovo">
                                                    <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                                                </li>
                                                <li runat="server" class="iconExcluir">
                                                    <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                                </li>
                                                <li runat="server" class="iconConsultar">
                                                    <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                                                </li>
                                                <li runat="server">
                                                    <asp:LinkButton ID="lnkConsultarSefaz" Text="Consultar Situação Sefaz" runat="server" class="iconConsultar" />
                                                </li>
                                                <li runat="server" class="iconLimpar">
                                                    <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                                                </li>
                                                <li runat="server" class="iconRelatorio">
                                                    <asp:LinkButton ID="lnkEspelho" Text="Espelho" runat="server" />
                                                </li>
                                                <li runat="server" class="iconRelatorio">
                                                    <asp:LinkButton ID="lnkExcel" Text="Excel" runat="server" />
                                                </li>
                                                <li runat="server" class="iconRelatorio">
                                                    <asp:LinkButton ID="lnkDACTE" Text="DACTE" runat="server" />
                                                </li>
                                                <li runat="server" class="iconMail">
                                                    <asp:LinkButton ID="lnkEnviarSEFAZ" Text="Enviar SEFAZ" runat="server" />
                                                </li>
                                                <li runat="server" class="iconMail">
                                                    <asp:LinkButton ID="lnkEnviarEmail" Text="Enviar E-mail" runat="server" />
                                                </li>
                                                <li runat="server" class="iconMais">
                                                    <asp:LinkButton ID="lnkDuplicar" Text="Duplicar" runat="server" />
                                                </li>
                                                <li runat="server" class="iconRelatorio">
                                                    <asp:UpdatePanel ID="updpnlVisualizar" runat="server">
                                                        <ContentTemplate>
                                                            <asp:LinkButton ID="lnkVisualizar" Text="Pré-Visualizar" runat="server" />
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:PostBackTrigger ControlID="lnkVisualizar" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </li>
                                                <li runat="server" class="iconMais">
                                                    <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                    <table style="width: 100%; padding: 0px;">
                                        <tr>
                                            <td valign="top">
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>EMPRESA</em></strong></legend>
                                                    <table style="border-top-style: none; border-right-style: none; border-left-style: none; border-bottom-style: none"
                                                        cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                                                <asp:Button ID="btnEmpresa" runat="server" Text=" > " UseSubmitBehavior="False" Enabled="False" />
                                                                <asp:Label ID="txtNomeDaEmpresa" runat="server" ForeColor="Blue" Font-Bold="False"
                                                                    Font-Names="Tahoma" Font-Italic="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEnderecoDaEmpresa" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtComplementoDaEmpresa" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">BAIRRO/DISTRITO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtBairroDaEmpresa" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">CEP</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCepDaEmpresa" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtInscricaoDaEmpresa" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                                    Font-Italic="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCnpjDaEmpresa" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                                    Font-Italic="False" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>TOMADOR</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                                                <asp:Label ID="txtNomeDoTomador" runat="server" ForeColor="Blue" Font-Bold="False"
                                                                    Font-Names="Tahoma" Font-Italic="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEnderecoDoTomador" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtComplementoDoTomador" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">BAIRRO/DISTRITO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtBairroDoTomador" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">CEP</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCepDoTomador" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtInscricaoDoTomador" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                                    Font-Italic="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCnpjDoTomador" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                                    Font-Italic="False" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <div id='divEmitente' runat="server" visible="false">
                                                    <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em; color: red;"><strong><em>EMITENTE</em></strong></legend>
                                                        <table cellspacing="0" cellpadding="0" width="100%">
                                                            <tr>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt; color: red;">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                                                    <asp:Label ID="txtNomeDoEmitente" runat="server" ForeColor="red" Font-Bold="False"
                                                                        Font-Names="Tahoma" Font-Italic="False" />
                                                                </td>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt; color: red;">ENDEREÇO</span></strong></label>&nbsp;<br />
                                                                    <asp:Label ID="txtEnderecoDoEmitente" runat="server" ForeColor="red" Font-Names="Tahoma" />
                                                                </td>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt; color: red;">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                                                    <asp:Label ID="txtComplementoDoEmitente" runat="server" ForeColor="red" Font-Names="Tahoma" />
                                                                </td>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt; color: red;">BAIRRO/DISTRITO</span></strong></label>&nbsp;<br />
                                                                    <asp:Label ID="txtBairroDoEmitente" runat="server" ForeColor="red" Font-Names="Tahoma" />
                                                                </td>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt; color: red;">CEP</span></strong></label>&nbsp;<br />
                                                                    <asp:Label ID="txtCepDoEmitente" runat="server" ForeColor="red" Font-Names="Tahoma" />
                                                                </td>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt; color: red;">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                                                    <asp:Label ID="txtInscricaoDoEmitente" runat="server" ForeColor="red" Font-Names="Tahoma"
                                                                        Font-Italic="False" />
                                                                </td>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt; color: red;">CNPJ</span></strong></label>&nbsp;<br />
                                                                    <asp:Label ID="txtCnpjDoEmitente" runat="server" ForeColor="red" Font-Names="Tahoma"
                                                                        Font-Italic="False" />
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </fieldset>
                                                </div>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>REMETENTE</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                                                <asp:Label ID="txtNomeDoRemetente" runat="server" ForeColor="Blue" Font-Bold="False"
                                                                    Font-Names="Tahoma" Font-Italic="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEnderecoDoRemetente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtComplementoDoRemetente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">BAIRRO/DISTRITO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtBairroDoRemetente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">CEP</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCepDoRemetente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtInscricaoDoRemetente" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                                    Font-Italic="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCnpjDoRemetente" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                                    Font-Italic="False" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>DESTINATÁRIO</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                                                <asp:Label ID="txtNomeDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />&nbsp;
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">CNPJ/CPF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCnpjDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                                    Font-Italic="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label><br />
                                                                <asp:Label ID="txtInscricaoDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td style="width: 115px">
                                                                <div id="divVencimentoFatura" runat="server">
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt">VENC. FATURA</span></strong></label>&nbsp;<br />
                                                                    <asp:TextBox ID="txtVencimentoFatura" runat="server" Width="60px" BorderStyle="None" CssClass="calendario" />
                                                                </div>
                                                            </td>
                                                            <td style="width: 115px">
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">DATA DA EMISSÃO</span></strong></label>&nbsp;<br />
                                                                <asp:TextBox ID="txtDataDeEmissao" runat="server" Width="60px" BorderStyle="None" CssClass="calendario" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;ENDEREÇO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEnderecoDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtComplementoDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">BAIRRO/DISTRITO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtBairroDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">CEP</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCepDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">MUNICÍPIO</span></strong></label><br />
                                                                <asp:Label ID="txtCidadeDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">UF</span></strong></label><br />
                                                                <asp:Label ID="txtEstadoDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">FONE/FAX</span></strong></label><br />
                                                                <asp:Label ID="txtTelefoneDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                                            </td>
                                                            <td style="width: 115px">
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">DATA DE SAÍDA</span></strong></label>&nbsp;<br />
                                                                <asp:TextBox ID="txtDataDeSaida" runat="server" Width="60px" BorderStyle="None" CssClass="calendario" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <ajaxToolkit:TabContainer ID="TcDados" runat="server" ActiveTabIndex="0" Width="100%">
                                                    <ajaxToolkit:TabPanel ID="TabNotasFiscais" runat="server">
                                                        <HeaderTemplate>
                                                            Notas Fiscais
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                            <div class="bordagrid" style="height: 150px;">

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
                                                                        <asp:BoundField DataField="TipoDocumento" HeaderText="Tipo">
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="CodigoOperacao" HeaderText="Operação">
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
                                                                        <asp:BoundField DataField="QuantidadeFisica" HeaderText="Físico" DataFormatString="{0:N0}">
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
                                                        </ContentTemplate>
                                                    </ajaxToolkit:TabPanel>
                                                    <ajaxToolkit:TabPanel ID="TabFaturasdeFrete" runat="server" HeaderText="Faturas">
                                                        <ContentTemplate>
                                                            <div class="bordagrid" style="height: 150px;">
                                                                <asp:GridView ID="grdFaturas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                                    ForeColor="#333333" GridLines="None" Width="100%">
                                                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <EditRowStyle BackColor="#999999" />
                                                                    <Columns>
                                                                        <asp:BoundField DataField="CodigoFatura" HeaderText="Fatura">
                                                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                                        </asp:BoundField>
                                                                        <asp:TemplateField HeaderText="Conveniado">
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("Conveniado.Nome")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento">
                                                                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                                                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="ValorDaFatura" DataFormatString="{0:n2}" HeaderText="Valor Fatura">
                                                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="ValorLancadoFatura" DataFormatString="{0:n2}" HeaderText="Valor Lan&#231;ado">
                                                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="ValorSaldoFatura" DataFormatString="{0:n2}" HeaderText="Saldo">
                                                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="CodigoTitulo" HeaderText="Titulo" Visible="False">
                                                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="CodigoEmpresa" HeaderText="Empresa" Visible="False">
                                                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="EndEmpresa" HeaderText="EndEmpresa" Visible="False">
                                                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="ConveniadoEnd" HeaderText="EndConveniado" Visible="False">
                                                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                                        </asp:BoundField>
                                                                    </Columns>
                                                                </asp:GridView>
                                                            </div>
                                                        </ContentTemplate>
                                                    </ajaxToolkit:TabPanel>
                                                    <ajaxToolkit:TabPanel ID="TabVencimentos" runat="server" HeaderText="Vencimentos">
                                                        <ContentTemplate>
                                                            <div class="bordagrid" style="height: 150px;">
                                                                <asp:GridView ID="grdVencimentos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                                    ForeColor="#333333" GridLines="None" Width="100%">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <Columns>
                                                                        <asp:TemplateField HeaderText="Fatura">
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("CodigoFatura")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Tipo">
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lblReceberPagar" runat="server" Text='<%#Eval("Titulo.ReceberPagar")%>' />
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Titulo">
                                                                            <ItemTemplate>
                                                                                <asp:HyperLink ID="hpTitulo" runat="server" Target="_blank" NavigateUrl='<%# Eval("CodigoCifrado", "WFTitulo.aspx?param={0}")%>' Text='<%# Eval("Titulo.Codigo")%>' />
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Sit. Financeira">
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lblSituacaoFinanceira" runat="server" Text='<%#Eval("Titulo.CodigoProvisao")%>' />
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Prorrogação">
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("Titulo.Prorrogacao", "{0:dd/MM/yyyy}")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Valor do Doc.">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("Titulo.ValorDoDocumento", "{0:N2}")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Descontos">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("Titulo.Descontos", "{0:N2}")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Deduções">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("Titulo.Deducoes", "{0:N2}")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Juros">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("Titulo.Juros", "{0:N2}")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Acréscimos">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("Titulo.Acrescimos", "{0:N2}")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Val. Líquido">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("Titulo.ValorLiquido", "{0:N2}")%>
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
                                                                <asp:GridView ID="grdVencimentosFNNovo" runat="server" AutoGenerateColumns="False"
                                                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <Columns>
                                                                        <asp:TemplateField HeaderText="Fatura">
                                                                            <ItemTemplate>
                                                                                <%#Eval("CodigoFatura")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Tipo">
                                                                            <ItemTemplate>
                                                                                <%#Eval("TituloNovo.ReceberPagar")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Titulo">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("TituloNovo.Codigo")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Sit. Financeira">
                                                                            <ItemTemplate>
                                                                                <%#Eval("TituloNovo.DescricaoProvisao")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Prorrogação">
                                                                            <ItemTemplate>
                                                                                <%#Eval("TituloNovo.Reprogramacao", "{0:dd/MM/yyyy}")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Valor do Doc.">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("TituloNovo.Valores.EncargoValorDocumento", "{0:N2}")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Descontos">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("TituloNovo.Valores.TotalDebitos", "{0:N2}")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Acréscimos">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("TituloNovo.Valores.TotalCreditos", "{0:N2}")%>
                                                                            </ItemTemplate>
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Val. Líquido">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                            <ItemTemplate>
                                                                                <%#Eval("TituloNovo.Valores.EncargoValorLiquido", "{0:N2}")%>
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
                                                        </ContentTemplate>
                                                    </ajaxToolkit:TabPanel>
                                                    <ajaxToolkit:TabPanel ID="TabContabilizacao" runat="server" HeaderText="Contabilização">
                                                        <ContentTemplate>
                                                            <div class="bordagrid" style="height: 150px;">
                                                                <asp:GridView ID="grdContabilizacao" runat="server" Width="100%" ForeColor="#333333" GridLines="None"
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
                                                                                <asp:Label ID="lblDescricao" runat="server" Text='<%# Eval("Conta.titulo")%>' />
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
                                                                        <asp:BoundField DataField="Historico" HeaderText="Hist&#243;rico">
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
                                                    <ajaxToolkit:TabPanel runat="server" ID="TabNFOrigem">
                                                        <HeaderTemplate>
                                                            NF de Origem
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                            <div class="menu_acoes">
                                                                <div class="acoes">
                                                                    <ul>
                                                                        <li class="iconMais" runat="server">
                                                                            <asp:LinkButton ID="lnkAddNota" Text="Adicionar" runat="server" />
                                                                        </li>
                                                                    </ul>
                                                                </div>
                                                            </div>
                                                            <div class="row" id="divPedidoReferenciaRPA" runat="server">
                                                                <div class="collbl">
                                                                    Período Inicial:
                                                                </div>
                                                                <div class="coltxt">
                                                                    <asp:TextBox ID="txtDataInicio" runat="server" CssClass="calendario" Width="88px"
                                                                        data-ToolTip="default" ToolTip="Data inicial da busca." />
                                                                </div>
                                                                <div class="collbl">
                                                                    Período Final:
                                                                </div>
                                                                <div class="coltxt">
                                                                    <asp:TextBox ID="txtDataFim" runat="server" CssClass="calendario" Width="88px"
                                                                        data-ToolTip="default" ToolTip="Data final da busca." />
                                                                </div>
                                                                <div class="collbl">
                                                                    Pedido:
                                                                </div>
                                                                <div class="coltxt">
                                                                    <asp:HiddenField ID="txtCodigoClienteNFReferencial" runat="server" />
                                                                    <asp:TextBox ID="txtPedidoNFReferencial" runat="server" Width="90px" CssClass="txtNumerico"
                                                                        Style="text-align: right" AutoPostBack="True" ToolTip="Caso o campo pedido esteja bloqueado é porque já existem notas referenciadas. Para liberá-lo será necessário excluí-las." />
                                                                </div>
                                                                <div class="coltxt">
                                                                    <asp:Button ID="btnPedidoNFReferencial" runat="server" Text=">" UseSubmitBehavior="False"
                                                                        CssClass="btn" />
                                                                </div>
                                                            </div>
                                                            <div class="bordagrid" style="height: 140px;">
                                                                <asp:GridView ID="grdNotasFretes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                                    ForeColor="#333333" GridLines="None" Width="100%">
                                                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                    <EditRowStyle BackColor="#999999" />
                                                                    <Columns>
                                                                        <asp:TemplateField HeaderText="Cliente">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="Label1" runat="server" Text='<%# Bind("DescClienteCompleto") %>' />
                                                                            </ItemTemplate>
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Depósito">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="Label2" runat="server" Text='<%# Bind("DescDepositoCompleto") %>' />
                                                                            </ItemTemplate>
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Destino">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="Label3" runat="server" Text='<%# Bind("DescDestinoCompleto") %>' />
                                                                            </ItemTemplate>
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Operação">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="Label4" runat="server" Text='<%# Bind("DescSubOperacaoCompleto") %>' />
                                                                            </ItemTemplate>
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:TemplateField>
                                                                        <asp:BoundField DataField="Codigo" HeaderText="Nota" HtmlEncode="False">
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento"
                                                                            HtmlEncode="False">
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:BoundField>
                                                                        <asp:TemplateField HeaderText="Tipo">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="Label5" runat="server" Text='<%# Bind("TipoDeDocumento.Descricao") %>' />
                                                                            </ItemTemplate>
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField>
                                                                            <ItemTemplate>
                                                                                <asp:ImageButton ID="imgExcluirNF" runat="server" ImageUrl="~/Images/deletar.gif"
                                                                                    Style="border: 0;" OnClick="imgExcluirNF_Click" data-ToolTip="default" ToolTip="Excluir"
                                                                                    OnClientClick="return confirm('Deseja realmente excluir o registro selecionado?');" />
                                                                            </ItemTemplate>
                                                                            <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                                            <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                                        </asp:TemplateField>
                                                                    </Columns>
                                                                </asp:GridView>
                                                                <asp:GridView runat="server" AutoGenerateColumns="False" CellPadding="4" GridLines="None"
                                                                    ForeColor="#333333" Width="100%" ID="grdNotasReferenciais" ShowFooter="True">
                                                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                                                                    <Columns>
                                                                        <asp:TemplateField HeaderText="Nota">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lblNota" runat="server" Text='<%# String.Format("{0:N0}", Eval("Nota_Id"))%>' />
                                                                            </ItemTemplate>
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:TemplateField>
                                                                        <asp:BoundField DataField="Serie_id" HeaderText="Serie"></asp:BoundField>
                                                                        <asp:BoundField DataField="Cliente_Id" HeaderText="Cliente"></asp:BoundField>
                                                                        <asp:BoundField DataField="EndCliente_id" HeaderText="End.">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <asp:TemplateField HeaderText="Nome">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lblClienteNFReferencial" runat="server" Text='<%# Bind("ParentOrigem.NotaFiscal.cliente.nome")%>' />
                                                                            </ItemTemplate>
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="Faz">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lblFazendaNFReferencial" runat="server" Text='<%# Bind("ParentOrigem.NotaFiscal.cliente.complemento")%>' />
                                                                            </ItemTemplate>
                                                                            <HeaderStyle HorizontalAlign="Left" />
                                                                            <ItemStyle HorizontalAlign="Left" />
                                                                        </asp:TemplateField>
                                                                        <asp:BoundField DataField="Quantidade" DataFormatString="{0:n2}" HeaderText="Qtd.">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <asp:BoundField DataField="Valor" DataFormatString="{0:n2}" HeaderText="Valor do Frete">
                                                                            <HeaderStyle HorizontalAlign="Right" />
                                                                            <ItemStyle HorizontalAlign="Right" />
                                                                        </asp:BoundField>
                                                                        <asp:TemplateField>
                                                                            <ItemTemplate>
                                                                                <asp:ImageButton ID="imgDelNFReferencial" runat="server" ImageUrl="~/Images/deletar.gif"
                                                                                    OnClick="imgDelNFReferencial_Click" OnClientClick="if(!(confirm('Confirma Exclusão da Nota Referencial?'))) return false;" />
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
                                                </ajaxToolkit:TabContainer>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>EMBARQUE / DESCARGA</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">LOCAL DE EMBARQUE/COLETA</span></strong></label><br />
                                                                <asp:Label ID="txtEmbarque" runat="server" ForeColor="Blue" Width="320px" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtMunicipioEmbarque" runat="server" ForeColor="Blue" Font-Bold="False" Width="200px" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEstadoEmbarque" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">LOCAL DE DESCARGA/DESTINO</span></strong></label><br />
                                                                <asp:Label ID="txtNomeOrigemDestino" runat="server" ForeColor="Blue" Font-Names="Tahoma" Width="320px" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtMunicipioOrigemDestino" runat="server" ForeColor="Blue" Font-Bold="False" Width="200px" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEstadoOrigemDestino" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>DADOS DO TRANSPORTADOR</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;RAZÃO SOCIAL</span></strong></label>&nbsp;<br />
                                                                <asp:Button ID="btnTransportador" runat="server" Text=" > " UseSubmitBehavior="False" />
                                                                <asp:Label ID="txtNomeDoTransportador" runat="server" Width="260px" ForeColor="Blue"
                                                                    Font-Bold="False" Font-Names="Tahoma" />&nbsp;&nbsp;
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;CÓDIGO ANTT</span></strong></label><span
                                                                        style="font-size: 5pt">&nbsp;<br />
                                                                    </span>&nbsp;<asp:Label ID="txtRNTRCTransportador" runat="server" Font-Names="Tahoma"
                                                                        ForeColor="Blue" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;ENDEREÇO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEnderecoDoTransportador" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">MUNICÍPIO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCidadeDoTransportador" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">UF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEstadoDoTransportador" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">CNPJ/CPF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCnpjDoTransportador" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtInscricaoDoTransportador" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>DADOS DO CAMINHÃO</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;PLACA 1</span></strong></label>&nbsp;<br />
                                                                <asp:Button ID="btnPlaca" runat="server" Text=" > " OnClick="btnPlaca_Click" UseSubmitBehavior="False" />
                                                                <asp:Label ID="txtCPlaca1" runat="server" Width="230px" ForeColor="Blue" Font-Bold="False"
                                                                    Font-Names="Tahoma" />&nbsp;
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCidadePlaca1" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEstadoPlaca1" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;RNTRC</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtRNTRCPlaca1" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;PROPRIETÁRIO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtProprietarioPlaca1" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;PLACA 2</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCPlaca2" runat="server" Width="260px" ForeColor="Blue" Font-Bold="False"
                                                                    Font-Names="Tahoma" />&nbsp;
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCidadePlaca2" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEstadoPlaca2" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;RNTRC</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtRNTRCPlaca2" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;PROPRIETÁRIO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtProprietarioPlaca2" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;PLACA 3</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCPlaca3" runat="server" Width="260px" ForeColor="Blue" Font-Bold="False"
                                                                    Font-Names="Tahoma" />&nbsp;
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCidadePlaca3" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEstadoPlaca3" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;RNTRC</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtRNTRCPlaca3" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;PROPRIETÁRIO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtProprietarioPlaca3" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>DADOS DO MOTORISTA</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;NOME</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtNomeDoMotorista" runat="server" Width="260px" ForeColor="Blue"
                                                                    Font-Bold="False" Font-Names="Tahoma" />&nbsp;
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;ENDEREÇO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEnderecoDoMotorista" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCidadeDoMotorista" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEstadoDoMotorista" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;CPF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCPFDoMotorista" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;HABILITAÇÃO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtHabilitacaoDoMotorista" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>FAVORECIDO DO CARTÃO</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;NOME</span></strong></label>&nbsp;<br />
                                                                <asp:Button ID="btnFavorecidoCartao" runat="server" Text=" > " UseSubmitBehavior="False"
                                                                    Enabled="False" />
                                                                <asp:Label ID="txtFavorecidoCartao" runat="server" Width="260px" ForeColor="Blue"
                                                                    Font-Bold="False" Font-Names="Tahoma" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;ENDEREÇO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEnderecoFavorecidoCartao" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCidadeFavorecidoCartao" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtEstadoFavorecidoCartao" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                            </td>
                                                            <td>
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt">&nbsp;CPF</span></strong></label>&nbsp;<br />
                                                                <asp:Label ID="txtCPFFavorecidoCartao" runat="server" ForeColor="Blue" Font-Bold="False" />
                                                                <asp:HiddenField ID="txtCodigoFavorecidoCartao" runat="server" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                            </td>
                                            <td style="width: 350px; align-content: center; vertical-align: top">
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>CHAVE DE ACESSO</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" border="0">
                                                        <tr>
                                                            <td valign="top" align="center">
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 6pt;">CHAVE DE ACESSO DO CT-e CONSULTA NO SITE: WWW.NFE.FAZENDA.GOV.BR</span></strong>
                                                                </label>
                                                                <br />
                                                                <asp:TextBox ID="txtChaveNFe" runat="server" Width="280px" Enabled="False" MaxLength="50"
                                                                    BorderColor="White" />
                                                                <asp:LinkButton ID="lnkVerificarChaveNFE" CssClass="lnk"
                                                                    data-tooltip="default" ToolTip="Consultar Chave NFE" runat="server" Text=" &gt; "
                                                                    OnClick="lnkVerificarChaveNFE_Click">
                                                                        <i class="fa fa-arrow-right seta"></i>
                                                                </asp:LinkButton>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <table>
                                                        <tr id="trFile" runat="server" visible="true">
                                                            <td style="width: 47px; white-space: nowrap; padding: 0;" valign="top">
                                                                <span class="HeaderSpanSecond">Arquivo:</span>
                                                            </td>
                                                            <td>
                                                                <uc:File ID="ucFile" runat="server" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>DADOS DO CT-e</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" border="0">
                                                        <tr>
                                                            <td style="font-size: 7pt; font-family: Tahoma" align="left">
                                                                <strong>NÚMERO:</strong>
                                                            </td>
                                                            <td style="border-right: #000000 1px solid; border-top: #000000 1px solid; font-size: 7pt; border-left: #000000 1px solid; width: 30px; border-bottom: #000000 1px solid"
                                                                valign="middle" align="center">
                                                                <asp:TextBox ID="txtNumero" runat="server" Width="160px" Style="color: blue; text-align: right"
                                                                    BorderColor="White" MaxLength="9" Font-Names="Tahoma" />
                                                            </td>
                                                            <td style="padding: 1px;" align="center">
                                                                <asp:Button ID="btnInutilizar" runat="server" CssClass="botao" Text="Inutilizar"
                                                                    UseSubmitBehavior="False" Width="90px"></asp:Button>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td style="font-size: 7pt; font-family: Tahoma" align="left">
                                                                <strong>SÉRIE:</strong>
                                                            </td>
                                                            <td style="border-right: #000000 1px solid; border-top: #000000 1px solid; font-size: 7pt; border-left: #000000 1px solid; width: 30px; border-bottom: #000000 1px solid"
                                                                valign="middle" align="center">
                                                                <asp:TextBox ID="txtSerie" runat="server" Width="160px" Style="color: blue; text-align: right"
                                                                    BorderColor="White" MaxLength="3" />
                                                            </td>
                                                            <td style="padding: 1px;" align="center">
                                                                <asp:Button ID="btnCadastro" runat="server" CssClass="botao" Text="Cadastro" UseSubmitBehavior="False"
                                                                    Width="90px" OnClick="btnCadastro_Click" Enabled="False"></asp:Button>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td style="font-size: 7pt; font-family: Tahoma" align="left">
                                                                <strong>SITUAÇÃO:</strong>
                                                            </td>
                                                            <td>
                                                                <asp:DropDownList ID="ddlSituacao" runat="server" Width="170px" Enabled="False">
                                                                </asp:DropDownList>
                                                            </td>
                                                            <td style="padding: 1px;" align="center">
                                                                <asp:Button ID="btnRelatorio" runat="server" CssClass="botao" Text="Relatório" UseSubmitBehavior="False"
                                                                    Width="90px" OnClick="btnRelatorio_Click" Enabled="False"></asp:Button>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td style="font-size: 7pt; font-family: Tahoma; width: 75px; white-space: nowrap;"
                                                                align="left">
                                                                <strong>PAGAMENTO:</strong>
                                                            </td>
                                                            <td>
                                                                <table cellpadding="0">
                                                                    <tr>
                                                                        <td style="border: 0;">
                                                                            <asp:RadioButton ID="rdoPamcard" runat="server" Text="PAMCARD" GroupName="grpConhecimento"
                                                                                AutoPostBack="True" />
                                                                        </td>
                                                                        <td style="border: 0;">
                                                                            <asp:RadioButton ID="rdoBanco" runat="server" Text="BANCO" Checked="True" GroupName="grpConhecimento"
                                                                                AutoPostBack="True" />
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </td>
                                                            <td style="padding: 1px;" align="center"></td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>INFORMAÇÕES E VALORES</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" border="0" width="97%">
                                                        <tr>
                                                            <td align="left" style="width: 115px; padding-left: 10px;">
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 10pt; color: black;">CARTÃO</span></strong></label>
                                                            </td>
                                                            <td align="right">
                                                                <asp:TextBox ID="txtCartao" runat="server" CssClass="txtCartao" Enabled="False" Width="195px"
                                                                    Font-Size="14pt" Style="text-align: right" BorderColor="White" MaxLength="20"
                                                                    Font-Names="Tahoma" Height="24px" AutoPostBack="True" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left" style="width: 115px; padding-left: 10px;">
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 10pt; color: black;">PESO</span></strong>
                                                                </label>
                                                            </td>
                                                            <td align="right">
                                                                <asp:TextBox ID="txtQuantidade" runat="server" BorderStyle="None" Width="195px" Height="24px"
                                                                    CssClass="numerico9" Font-Size="14pt" Style="text-align: right;" Enabled="False" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left" style="width: 115px; padding-left: 10px;">
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 10pt; color: red;">TARIFA (TON)</span></strong>
                                                                </label>
                                                            </td>
                                                            <td align="right">
                                                                <asp:TextBox ID="txtUnitario" runat="server" BorderStyle="None" Width="195px" Height="24px"
                                                                    AutoPostBack="True" CssClass="txtDecimal4" OnTextChanged="txtUnitario_TextChanged"
                                                                    Font-Size="14pt" Style="color: red; text-align: right;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left" style="width: 115px; padding-left: 10px;">
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 10pt; color: black;">TOTAL</span></strong>
                                                                </label>
                                                            </td>
                                                            <td align="right">
                                                                <asp:Label ID="txtValor" runat="server" Font-Bold="False" Font-Size="14pt" Font-Names="Tahoma"
                                                                    Height="24px" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left" style="width: 115px; padding-left: 10px;">
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 10pt; color: black;">OPERAÇÃO</span></strong></label>
                                                            </td>
                                                            <td align="right">
                                                                <asp:Label ID="txtOperacao" runat="server" Font-Bold="False" Font-Size="7pt" Font-Names="Tahoma"
                                                                    Height="36px" Style="padding-right: 5px; vertical-align: middle;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left" style="width: 115px">&nbsp;</td>
                                                            <td align="right">
                                                                <asp:Label ID="txtEntSai" runat="server" Font-Bold="False" Font-Size="7pt" Font-Names="Tahoma" Style="padding-right: 5px; vertical-align: middle;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left" style="width: 115px; padding-left: 10px;">
                                                                <label class="titulo">
                                                                    <strong><span style="font-size: 10pt; color: black;">CFOP</span></strong></label>
                                                            </td>
                                                            <td align="right">
                                                                <asp:Label ID="txtCfop" runat="server" Font-Bold="False" Font-Size="7pt" Font-Names="Tahoma"
                                                                    Height="24px" Style="padding-right: 5px;" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>ENCARGOS</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" border="0" width="100%">
                                                        <tr>
                                                            <td>
                                                                <asp:Panel ID="pnlEncargos" runat="server" CssClass="bordagrid" Height="150px"
                                                                    ScrollBars="Vertical" Width="100%">
                                                                    <asp:GridView ID="gridEncargos" runat="server" CellPadding="3" ForeColor="#333333"
                                                                        GridLines="None" Width="100%" AutoGenerateColumns="False">
                                                                        <EditRowStyle BackColor="#999999" />
                                                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                        <Columns>
                                                                            <asp:BoundField HeaderText="Descri&#231;&#227;o" DataField="Codigo">
                                                                                <HeaderStyle HorizontalAlign="Left" />
                                                                                <ItemStyle HorizontalAlign="Left" />
                                                                            </asp:BoundField>
                                                                            <asp:TemplateField HeaderText="Valor">
                                                                                <ItemTemplate>
                                                                                    <asp:TextBox ID="txtValorEncargo" Style="text-align: right" CssClass="txtDecimal"
                                                                                        Text='<%# Eval("Valor") %>' runat="server" BorderColor="White" Width="60px" Enabled="false" />
                                                                                </ItemTemplate>
                                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                                <ItemStyle HorizontalAlign="Right" />
                                                                            </asp:TemplateField>
                                                                            <asp:TemplateField>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lblSinal" runat="server" Text='<%# Eval("Sinal") %>' />
                                                                                </ItemTemplate>
                                                                                <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                                                                <ItemStyle HorizontalAlign="Center" Width="30px" />
                                                                            </asp:TemplateField>
                                                                        </Columns>
                                                                    </asp:GridView>
                                                                </asp:Panel>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="right" style="padding-top: 5px;">
                                                                <asp:Button ID="btnRecontabilizar" runat="server" CssClass="botao" Text="Recontabilizar" />
                                                                <asp:Button ID="btnCalcular" runat="server" CssClass="botao" Text="Calcular" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                                <fieldset style="border: 1px solid #BBBBBB;">
                                                    <legend style="font-size: 0.8em;"><strong><em>OBSERVAÇÃO</em></strong></legend>
                                                    <table cellspacing="0" cellpadding="0" border="0" width="100%">
                                                        <tr>
                                                            <td class="subtitulodiv" style="text-align: center;">
                                                                <label>
                                                                    Observação
                                                                </label>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td style="padding: 1px;">
                                                                <asp:TextBox ID="txtObservacao" runat="server" Height="155px" TextMode="MultiLine"
                                                                    Width="100%" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </fieldset>
                                            </td>
                                        </tr>
                                    </table>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                        </ajaxToolkit:TabContainer>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:NFOrigem ID="ucNFOrigem" runat="server" />
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:ConsultaPlacas ID="ucConsultaPlacas" runat="server" />
    <uc:ConsultaEstados ID="ucConsultaEstados" runat="server" />
    <uc:ConsultaCodMunicipios ID="ucConsultaCodMunicipios" runat="server" />
    <uc:ConsultaCTeXNotas ID="ucConsultaCTeXNotas" runat="server" />
    <uc:ConsultaRelatorio ID="ucConsultaRelatorio" runat="server" />
    <uc:ConsultaCadastro ID="ucConsultaCadastro" runat="server" />
    <uc:Inutilizacao ID="ucInutilizacao" runat="server" />
    <uc:EmailNFe ID="ucEmailNFe" runat="server" />
    <uc:EmitirCTe ID="ucEmitirCTe" runat="server" />
</asp:Content>
