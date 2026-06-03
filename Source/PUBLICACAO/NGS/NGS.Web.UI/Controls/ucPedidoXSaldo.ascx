<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucPedidoXSaldo.ascx.vb"
    Inherits="NGS.Web.UI.ucPedidoXSaldo" %>
<script type="text/javascript">
    function pageLoadPedidoXSaldo() {
        $(".calendario").setMask('date');

        $("div.accordion").accordion({
            active: false,
            collapsible: true,
            alwaysOpen: false,
            heightStyle: "content",
            autoHeight: false,
            clearStyle: true
        });
    }
    $(document).ready(function () {
        pageLoadPedidoXSaldo();
    });

    var prmPedidoXSaldo = Sys.WebForms.PageRequestManager.getInstance();
    prmPedidoXSaldo.add_endRequest(pageLoadPedidoXSaldo);
</script>

<div id="divPedidoXSaldo" class="uc" title="Consulta de Pedido x Saldo" style="display: none;">
    <asp:UpdatePanel ID="updpnlPedidoXSaldo" runat="server">
        <ContentTemplate>
            <script type="text/javascript">
                function selectAllCheckboxes(chkAll) {
                    var chk = $('#' + chkAll.id);
                    var checked = chk.attr('checked') == "checked";
                    $("input[type='checkbox']", "#MainContent_ucPedidoxSaldo_gridGlobalDiretaItens").not("#chkAll").each(function () {
                        $(this).attr("checked", checked);
                    });
                }

                function selectAllCheckboxesFix(chkAll) {
                    var chk = $('#' + chkAll.id);
                    var checked = chk.attr('checked') == "checked";
                    $("input[type='checkbox']", "#MainContent_ucPedidoxSaldo_GridAFixarItens").not("#chkAll").each(function () {
                        $(this).attr("checked", checked);
                    });
                }
            </script>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="primeiraVez" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton ID="lnkConsultarPedidos" CssClass="iconConsultar" Text="Consultar pedidos"
                                runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton ID="lnkConsultarFixacoes" CssClass="iconConsultar" Text="Consultar fixações"
                                runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton ID="lnkDevolucaoVNotas" CssClass="iconConsultar" Text="Devolução vários Pedidos"
                                runat="server" Visible="false" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton ID="lnkFechar" CssClass="iconSair" runat="server" Text="Fechar" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc" style="width: 128px;">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeEmpresa" runat="server" Width="500px" />
                    <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnEmpresa" runat="server" OnClick="btnEmpresa_Click" Text="&gt;"
                        CssClass="btn" />
                </div>
                <div class="collbluc" style="width: 128px;">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdTodos" runat="server" GroupName="situacao" Text="Todos" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdComSaldo" runat="server" Checked="True" GroupName="situacao"
                        Text="Com Saldo" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdSemSaldo" runat="server" GroupName="situacao" Text="Sem Saldo" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc" style="width: 128px;">
                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente." Enabled="false" runat="server" Text="Consolidar Cliente:"
                        data-ToolTip="default" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeCliente" runat="server" Width="500px" />
                    <asp:HiddenField ID="txtCodCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" OnClick="btnCliente_Click" Text="&gt;"
                        CssClass="btn" />
                </div>
                <div class="collbluc" style="width: 128px;">
                    Fiscal:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdFTodos" runat="server" Checked="True" GroupName="fiscal" Text="Todos" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdFAberto" runat="server" GroupName="fiscal" Text="Aberto" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdFFechado" runat="server" GroupName="fiscal" Text="Fechado" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc" style="width: 128px;">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlSafra_SelectedIndexChanged"
                        Width="225px" />
                </div>
                <div class="collbluc" style="width: 128px;">
                    Pedido:
                </div>
                <div class="coltxt" style="width: 171px;">
                    <asp:TextBox ID="txtPedido" runat="server" ClientIDMode="Static" CssClass="txtNumerico"
                        Width="131px" />
                </div>
                <div class="collbluc" style="width: 128px;">
                    <asp:CheckBox ID="chkPeriodo" runat="server" AutoPostBack="True" Checked="True" Text="Usar Período:" />
                </div>
                <div class="coltxt">
                    <asp:Panel ID="pnlPeriodo" runat="server">
                        <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" EnableTheming="True"
                            Width="80px" />
                        <label>
                            à
                        </label>
                        <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="80px" />
                    </asp:Panel>
                </div>
            </div>
            <asp:Panel ID="PnlSelecao" runat="server" Visible="false">
                <div id="divPedidosGlobalDireta" class="accordion" runat="server">
                    <h3>
                        <label>
                            Pedidos Com Entrega Futura / Entrega Direta
                        </label>
                    </h3>
                    <div>
                        <div class="bordagrid" style="height: 200px;">
                            <asp:GridView ID="gridGlobalDireta" runat="server" CellPadding="0" ForeColor="#333333"
                                GridLines="None" Width="100%" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                                    <asp:BoundField DataField="DescricaoCliente" HeaderText="Cliente" ItemStyle-HorizontalAlign="Left"
                                        HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="CodigoPedido" HeaderText="Pedido" ItemStyle-HorizontalAlign="Left"
                                        HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="PedidoEfetivo" HeaderText="P.Efetivo" ItemStyle-HorizontalAlign="Left"
                                        HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="DataPedido" DataFormatString="{0:dd-MM-yyyy}" HeaderText="Pedido de"
                                        ItemStyle-HorizontalAlign="Center" HeaderStyle-HorizontalAlign="Center" />
                                    <asp:BoundField DataField="DescricaoSuboperacao" HeaderText="Operação" ItemStyle-HorizontalAlign="Left"
                                        HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="DescricaoMoeda" HeaderText="Moeda" HeaderStyle-HorizontalAlign="Left"
                                        ItemStyle-HorizontalAlign="Left" />
                                    <asp:TemplateField HeaderText="Qtde Pedido">
                                        <ItemTemplate>
                                            <div style="float: Left; margin-right: 3px;">
                                                <asp:Literal ID="LPROG" runat="server" Text="Programada" />
                                                <br />
                                                <asp:Literal ID="LCONT" runat="server" Text="Contratada" />
                                                <br />
                                                <asp:Literal ID="LCOM" runat="server" Text="Comercial" />
                                            </div>
                                            <div style="float: right;">
                                                <asp:Literal ID="Literal4" runat="server" Text='<%# String.Format("{0:N2}",eval("qtdeProgramada")) %>' />
                                                <br />
                                                <asp:Literal ID="Literal5" runat="server" Text='<%# String.Format("{0:N2}",eval("qtdecontratadofiscal")) %>' />
                                                <br />
                                                <asp:Literal ID="Literal6" runat="server" Text='<%# String.Format("{0:N2}",eval("QtdeProgramadaComercializacao")) %>' />
                                            </div>
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Right" BackColor="#FFFFCC" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Vlrs Pedido">
                                        <ItemTemplate>
                                            <div style="float: right;">
                                                <asp:Literal ID="L01" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                                <br />
                                                <asp:Literal ID="L02" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                            </div>
                                            <div style="float: right; margin-right: 3px;">
                                                <asp:Literal ID="L03" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrPedidoOficial")) %>' />
                                                <br />
                                                <asp:Literal ID="L04" runat="server" Text='<%# String.Format("{0:N2}",eval("vlrPedidoMoeda")) %>' />
                                            </div>
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Right" BackColor="#FFFFCC" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Qtde Entregue">
                                        <ItemTemplate>
                                            <div style="float: right;">
                                                <asp:Label ID="L17" runat="server" Text='<%# String.Format("{0:N2}",eval("QtdeEntregueFiscalDireta")) %>' />
                                                <br />
                                                <asp:Label ID="L18" runat="server" Text='<%# String.Format("{0:N2}",eval("QtdeEntregueFiscalGlobal")) %>' />
                                                <br />
                                                <asp:Label ID="L19" runat="server" Text='<%# String.Format("{0:N2}",eval("QtdeEntregueFiscalRemessa")) %>' />
                                            </div>
                                            <div style="float: right; margin-right: 3px;">
                                                <asp:Literal ID="LD" runat="server" Text="Direta" />
                                                <br />
                                                <asp:Literal ID="LG" runat="server" Text="Futura" />
                                                <br />
                                                <asp:Literal ID="LR" runat="server" Text="Remessa" />
                                            </div>
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" BackColor="#99CCFF" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Vlr Entregue">
                                        <ItemTemplate>
                                            <div runat="server" style="float: right;">
                                                <asp:Literal ID="L08CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                                <br />
                                                <asp:Literal ID="L09CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                                <br />
                                                <asp:Literal ID="L10CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                            </div>
                                            <div runat="server" style="float: right;">
                                                <asp:Literal ID="L08" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaMoedaDiretaBruto")) %>' />
                                                <br />
                                                <asp:Literal ID="L09" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaMoedaGlobalBruto")) %>' />
                                                <br />
                                                <asp:Literal ID="L10" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaMoedaRemessaBruto")) %>' />
                                            </div>
                                            <div style="float: right;">
                                                <asp:Literal ID="L05CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                                <br />
                                                <asp:Literal ID="L06CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                                <br />
                                                <asp:Literal ID="L07CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                            </div>
                                            <div style="float: right;">
                                                <asp:Literal ID="L05" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaOficialDiretaBruto")) %>' />
                                                <br />
                                                <asp:Literal ID="L06" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaOficialGlobalBruto")) %>' />
                                                <br />
                                                <asp:Literal ID="L07" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaOficialRemessaBruto")) %>' />
                                            </div>
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" BackColor="#99CCFF" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Qtde Saldo">
                                        <ItemTemplate>
                                            <div style="float: right;">
                                                <asp:Literal ID="L20" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoQtdeDiretoFiscal")) %>' />
                                                <br />
                                                <asp:Literal ID="L21" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoQtdeGlobalFiscal")) %>' />
                                                <br />
                                                <asp:Literal ID="L22" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoQtdeRemessaFiscal")) %>' />
                                            </div>
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" BackColor="#FFFFCC" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Vlr Saldo">
                                        <ItemTemplate>
                                            <div runat="server" style="float: right;">
                                                <asp:Literal ID="L14CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                                <br />
                                                <asp:Literal ID="L15CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                                <br />
                                                <asp:Literal ID="L16CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                            </div>
                                            <div runat="server" style="float: right;">
                                                <asp:Literal ID="L14" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorMoedaGlobalDireto")) %>' />
                                                <br />
                                                <asp:Literal ID="L15" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorMoedaGlobalDireto")) %>' />
                                                <br />
                                                <asp:Literal ID="L16" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorMoedaRemessa")) %>' />
                                            </div>
                                            <div style="float: right;">
                                                <asp:Literal ID="L11CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                                <br />
                                                <asp:Literal ID="L12CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                                <br />
                                                <asp:Literal ID="L13CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                            </div>
                                            <div style="float: right;">
                                                <asp:Literal ID="L11" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorOficialGlobalDireto")) %>' />
                                                <br />
                                                <asp:Literal ID="L12" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorOficialGlobalDireto")) %>' />
                                                <br />
                                                <asp:Literal ID="L13" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorOficialRemessa")) %>' />
                                            </div>
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" BackColor="#FFFFCC" />
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgAbertoFechado" runat="server" Height="22px" ImageUrl="~/images/certo.jpg"
                                                data-ToolTip="default" ToolTip="Fiscal Aberto" Width="22px" Style="border: 0;" />
                                            <asp:HiddenField ID="hidFiscalAberto" runat="server" Value='<%# eval("FiscalAberto") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </div>
                </div>
                <div id="divPedidosDeposito" class="accordion" runat="server">
                    <h3>
                        <label>
                            Pedidos de Deposito
                        </label>
                    </h3>
                    <div>
                        <div class="bordagrid" style="height: 200px;">
                            <asp:GridView ID="gridDeposito" runat="server" CellPadding="0" ForeColor="#333333"
                                GridLines="None" Width="100%" PageSize="5" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                                    <asp:BoundField DataField="DescricaoCliente" HeaderText="Cliente" ItemStyle-HorizontalAlign="Left"
                                        HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="CodigoPedido" HeaderText="Pedido" ItemStyle-HorizontalAlign="Right"
                                        HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="DescricaoSuboperacao" HeaderText="Operação" ItemStyle-HorizontalAlign="Left"
                                        HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="DataPedido" DataFormatString="{0:dd-MM-yyyy}" HeaderText="Pedido de"
                                        ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="DescricaoMoeda" HeaderText="Moeda" ItemStyle-HorizontalAlign="Left"
                                        HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="QtdeProgramada" DataFormatString="{0:N2}" HeaderText="Qtde Programada"
                                        ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="QtdeContratadoFiscal" DataFormatString="{0:N2}" HeaderText="Qtde Contratada"
                                        ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="QtdeEntregueFiscalDeposito" DataFormatString="{0:N2}"
                                        HeaderText="Qtde Entregue" ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="VlrNotaOficialDepositoBruto" DataFormatString="{0:N2}"
                                        HeaderText="Valor Entregue" ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgAbertoFechado" runat="server" Height="22px" ImageUrl="~/images/certo.jpg"
                                                data-ToolTip="default" ToolTip="Fiscal Aberto" Width="22px" Style="border: 0;" />
                                            <asp:HiddenField ID="hidFiscalAberto" runat="server" Value='<%# eval("FiscalAberto") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </div>
                </div>
                <div id="divPedidosAFixar" class="accordion" runat="server">
                    <h3>
                        <label>
                            Pedidos de Fixação
                        </label>
                    </h3>
                    <div>
                        <div class="bordagrid" style="height: 200px;">
                            <asp:GridView ID="GridAFixar" runat="server" CellPadding="0" ForeColor="#333333"
                                GridLines="None" Width="100%" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                                    <asp:BoundField DataField="DescricaoCliente" HeaderText="Cliente" ItemStyle-HorizontalAlign="Left"
                                        HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="CodigoPedido" HeaderText="Pedido" ItemStyle-HorizontalAlign="Right"
                                        HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="DescricaoSuboperacao" HeaderText="Operação" ItemStyle-HorizontalAlign="Left"
                                        HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="DataPedido" DataFormatString="{0:dd-MM-yyyy}" HeaderText="Pedido de"
                                        ItemStyle-HorizontalAlign="Left" HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="DescricaoMoeda" HeaderText="Moeda" ItemStyle-HorizontalAlign="Left"
                                        HeaderStyle-HorizontalAlign="Left" />
                                    <asp:BoundField DataField="QtdeProgramada" DataFormatString="{0:N2}" HeaderText="Qtde Programada"
                                        ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="QtdeContratadoFiscal" DataFormatString="{0:N2}" HeaderText="Qtde Contratado"
                                        ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="QtdeEntregueFiscalAFixar" DataFormatString="{0:N2}" HeaderText="Qtde Entregue"
                                        ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="VlrNotaOficialAfixarBruto" DataFormatString="{0:N2}" HeaderText="Vlr Entregue R$"
                                        ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="VlrNotaMoedaAfixarBruto" DataFormatString="{0:N2}" HeaderText="Vlr Entregue $"
                                        ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="QtdeFixacao" DataFormatString="{0:N2}" HeaderText="Qtde Fixada"
                                        ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="VlrFixacaoOficial" DataFormatString="{0:N2}" HeaderText="Vlr Fixado R$"
                                        ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:BoundField DataField="SaldoQtdeAFixar" DataFormatString="{0:N2}" HeaderText="Saldo Qtde AFixar"
                                        ItemStyle-HorizontalAlign="Right" HeaderStyle-HorizontalAlign="Right" />
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgAbertoFechado" runat="server" Height="22px" ImageUrl="~/images/certo.jpg"
                                                data-ToolTip="default" ToolTip="Fiscal Aberto" Width="22px" Style="border: 0;" />
                                            <asp:HiddenField ID="hidFiscalAberto" runat="server" Value='<%# eval("FiscalAberto") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </div>
                </div>
                <div id="divProdutosDoPedido" runat="server" visible="false">
                    <div class="subtitulodiv">
                        Produtos do Pedido
                    </div>
                    <div class="row">
                        <div class="collbluc" style="width: 128px;">
                            Operação Comercial:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlOperacao" runat="server" Width="500px" AutoPostBack="True"
                                OnSelectedIndexChanged="ddlOperacao_SelectedIndexChanged" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkTodasOperacoes" runat="server" AutoPostBack="true" Text="Listar Todas" />
                        </div>
                        <div class="coltxt" style="float: right; font-size: 24px; font-weight: bold; width: 170px; text-align: center;">
                            <label>
                                PEDIDO
                            </label>
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc" style="width: 128px;">
                            Filial:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlFilial" runat="server" Width="600px" AutoPostBack="True" />
                        </div>
                        <div class="coltxt" style="float: right; font-size: 24px; font-weight: bold; width: 170px; text-align: center;">
                            <asp:Label ID="lblPedidoSelecionado" runat="server" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc" style="width: 128px;">
                            Data Nota :
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtDtNota" runat="server" CssClass="calendario" Width="90px" />
                        </div>
                        <div class="collbluc" style="width: 128px;">
                            Data Saída/Entrada:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtDtMovimento" runat="server" CssClass="calendario" Width="90px" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkNossaEmissao" runat="server" AutoPostBack="True" Font-Bold="True"
                                OnCheckedChanged="chkNossaEmissao_CheckedChanged" Text="Nossa Emissão" />
                        </div>
                        <asp:Panel ID="pnlNossaEmissao" runat="server">
                            <div class="collbluc" style="width: 128px;">
                                <asp:Label ID="lblNumNota" runat="server" Text="Num. Nota" />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNota" CssClass="txtNumerico" runat="server" Width="80px" MaxLength="9" />
                            </div>
                            <div class="collbluc" style="width: 128px;">
                                <asp:Label ID="lblSerieNota" runat="server" Text="Série Nota" />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSerie" runat="server" Width="50px" />
                            </div>
                        </asp:Panel>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkNotaRetroativa" runat="server" Text="Nota Retroativa" Visible="false" />
                        </div>
                        <div class="coltxt" style="float: right; font-size: 12px; font-weight: bold;">
                            <asp:Label ID="lblSafraSelecionada" runat="server" Text="" />
                        </div>
                    </div>
                    <div class="bordagrid" style="height: 160px;">
                        <asp:GridView ID="gridGlobalDiretaItens" runat="server" AutoGenerateColumns="False"
                            CellPadding="0" ForeColor="#333333" GridLines="None" Width="100%">
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#999999" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                            <Columns>
                                <asp:TemplateField>
                                    <HeaderTemplate>
                                        <input id="chkAll" onclick="selectAllCheckboxes(this);" type="checkbox" />
                                    </HeaderTemplate>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chkSelecionado" runat="server" AutoPostBack="True" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="NomeProduto" HeaderText="Nome" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="Produto Entrada">
                                    <ItemTemplate>
                                        <asp:DropDownList ID="ddlXmlProdutoXDePara" runat="server" />
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="Lote" HeaderText="Lote" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Classificacao" HeaderText="Class." HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="DescricaoEmbalagem" HeaderText="Embalagem" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="Qtde Pedido">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" BackColor="#FFFFCC" />
                                    <ItemTemplate>
                                        <div style="float: right;">
                                            <asp:Literal ID="Literal2" runat="server" Text='<%# eval("Unidade") %>' />
                                            <br />
                                            <asp:Literal ID="Literal1" runat="server" Text='<%# eval("Unidade") %>' />
                                            <br />
                                            <asp:Literal ID="Literal8" runat="server" Text='<%# eval("UnidadeComercializacao") %>' />
                                        </div>
                                        <div style="float: right;">
                                            <asp:Literal ID="Literal6" runat="server" Text='<%# String.Format("{0:N2}",eval("qtdeProgramada")) %>' />
                                            <br />
                                            <asp:Literal ID="Literal7" runat="server" Text='<%# String.Format("{0:N2}",eval("qtdecontratadofiscal")) %>' />
                                            <br />
                                            <asp:Literal ID="Literal9" runat="server" Text='<%# String.Format("{0:N2}",eval("QtdeProgramadaComercializacao")) %>' />
                                        </div>
                                        <div style="float: right;">
                                            <asp:Literal ID="LPROG" runat="server" Text="Programada" />
                                            <br />
                                            <asp:Literal ID="LCONT" runat="server" Text="Contratada" />
                                            <br />
                                            <asp:Literal ID="LCOM" runat="server" Text="Comercial" />
                                        </div>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Vlrs Pedido">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" BackColor="#FFFFCC" />
                                    <ItemTemplate>
                                        <div style="float: right;">
                                            <asp:Literal ID="L23" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                            <br />
                                            <asp:Literal ID="L25" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                        </div>
                                        <div style="float: right;">
                                            <asp:Literal ID="L24" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrPedidoOficial")) %>' />
                                            <br />
                                            <asp:Literal ID="L26" runat="server" Text='<%# String.Format("{0:N2}",eval("vlrPedidoMoeda")) %>' />
                                        </div>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Qtde Entregue">
                                    <ItemTemplate>
                                        <div style="float: right;">
                                            <asp:Literal ID="L27UN" runat="server" Text='<%# eval("Unidade") %>' />
                                            <br />
                                            <asp:Literal ID="L28UN" runat="server" Text='<%# eval("Unidade") %>' />
                                            <br />
                                            <asp:Literal ID="L29UN" runat="server" Text='<%# eval("Unidade") %>' />
                                        </div>
                                        <div style="float: right;">
                                            <asp:Label ID="L27" runat="server" Text='<%# String.Format("{0:N2}",eval("QtdeEntregueFiscalDireta")) %>' />
                                            <br />
                                            <asp:Label ID="L28" runat="server" Text='<%# String.Format("{0:N2}",eval("QtdeEntregueFiscalGlobal")) %>' />
                                            <br />
                                            <asp:Label ID="L29" runat="server" Text='<%# String.Format("{0:N2}",eval("QtdeEntregueFiscalRemessa")) %>' />
                                        </div>
                                        <div style="float: right;">
                                            <asp:Literal ID="LD0" runat="server" Text="Direta" />
                                            <br />
                                            <asp:Literal ID="LG0" runat="server" Text="Futura" />
                                            <br />
                                            <asp:Literal ID="LR0" runat="server" Text="Remessa" />
                                        </div>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" BackColor="#99CCFF" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Vlr Entregue">
                                    <ItemTemplate>
                                        <div runat="server" style="float: right;">
                                            <asp:Literal ID="L34CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                            <br />
                                            <asp:Literal ID="L36CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                            <br />
                                            <asp:Literal ID="L38CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                        </div>
                                        <div runat="server" style="float: right;">
                                            <asp:Literal ID="L34" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaMoedaDiretaBruto")) %>' />
                                            <br />
                                            <asp:Literal ID="L36" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaMoedaGlobalBruto")) %>' />
                                            <br />
                                            <asp:Literal ID="L38" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaMoedaRemessaBruto")) %>' />
                                        </div>
                                        <div style="float: right;">
                                            <asp:Literal ID="L33CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                            <br />
                                            <asp:Literal ID="L35CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                            <br />
                                            <asp:Literal ID="L37CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                        </div>
                                        <div style="float: right;">
                                            <asp:Literal ID="L33" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaOficialDiretaBruto")) %>' />
                                            <br />
                                            <asp:Literal ID="L35" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaOficialGlobalBruto")) %>' />
                                            <br />
                                            <asp:Literal ID="L37" runat="server" Text='<%# String.Format("{0:N2}",eval("VlrNotaOficialRemessaBruto")) %>' />
                                        </div>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" BackColor="#99CCFF" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Qtde Saldo">
                                    <ItemTemplate>
                                        <div style="float: right;">
                                            <asp:Literal ID="L30UN" runat="server" Text='<%# eval("Unidade") %>' />
                                            <br />
                                            <asp:Literal ID="L31UN" runat="server" Text='<%# eval("Unidade") %>' />
                                            <br />
                                            <asp:Literal ID="L32UN" runat="server" Text='<%# eval("Unidade") %>' />
                                        </div>
                                        <div style="float: right;">
                                            <asp:Literal ID="L30" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoQtdeDiretoFiscal")) %>' />
                                            <br />
                                            <asp:Literal ID="L31" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoQtdeGlobalFiscal")) %>' />
                                            <br />
                                            <asp:Literal ID="L32" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoQtdeRemessaFiscal")) %>' />
                                        </div>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" BackColor="#FFFFCC" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Vlr Saldo">
                                    <ItemTemplate>
                                        <div runat="server" style="float: right;">
                                            <asp:Literal ID="L40CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                            <br />
                                            <asp:Literal ID="L42CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                            <br />
                                            <asp:Literal ID="L44CF" runat="server" Text='<%# eval("CifraoPedido") %>' />
                                        </div>
                                        <div runat="server" style="float: right;">
                                            <asp:Literal ID="L40" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorMoedaGlobalDireto")) %>' />
                                            <br />
                                            <asp:Literal ID="L42" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorMoedaGlobalDireto")) %>' />
                                            <br />
                                            <asp:Literal ID="L44" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorMoedaRemessa")) %>' />
                                        </div>
                                        <div style="float: right;">
                                            <asp:Literal ID="L39CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                            <br />
                                            <asp:Literal ID="L41CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                            <br />
                                            <asp:Literal ID="L43CF" runat="server" Text='<%# eval("CifraoOficial") %>' />
                                        </div>
                                        <div style="float: right;">
                                            <asp:Literal ID="L39" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorOficialGlobalDireto")) %>' />
                                            <br />
                                            <asp:Literal ID="L41" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorOficialGlobalDireto")) %>' />
                                            <br />
                                            <asp:Literal ID="L43" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoValorOficialRemessa")) %>' />
                                        </div>
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" BackColor="#FFFFCC" />
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                        <asp:GridView ID="gridDepositoItens" runat="server" AutoGenerateColumns="False" CellPadding="0"
                            ForeColor="#333333" GridLines="None" Width="100%">
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#999999" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chkSelecionado" runat="server" />
                                    </ItemTemplate>
                                    <HeaderTemplate>
                                        <asp:CheckBox ID="chkAll" runat="server" />
                                    </HeaderTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="NomeProduto" HeaderText="Nome" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="Produto Entrada">
                                    <ItemTemplate>
                                        <asp:DropDownList ID="ddlXmlProdutoXDePara" runat="server" />
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="Unidade" HeaderText="Un." HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Lote" HeaderText="Lote" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Classificacao" HeaderText="Class." HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="DescricaoEmbalagem" HeaderText="Embalagem" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="QtdeProgramada" DataFormatString="{0:N2}" HeaderText="Qtde Programada"
                                    HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="QtdeContratadoFiscal" DataFormatString="{0:N2}" HeaderText="Qtde Contratada"
                                    HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="QtdeEntregueFiscalDeposito" DataFormatString="{0:N2}"
                                    HeaderText="Qtde Entregue" HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="VlrNotaOficialDepositoBruto" DataFormatString="{0:N2}"
                                    HeaderText="Valor Entregue" HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                            </Columns>
                        </asp:GridView>
                        <asp:GridView ID="GridAFixarItens" runat="server" AutoGenerateColumns="False" CellPadding="0"
                            ForeColor="#333333" GridLines="None" Width="100%">
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#999999" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chkSelecionado" runat="server" />
                                    </ItemTemplate>
                                    <HeaderTemplate>
                                        <asp:CheckBox ID="chkAll" onclick="selectAllCheckboxesFix(this);" runat="server" />
                                    </HeaderTemplate>
                                    <HeaderStyle HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="NomeProduto" HeaderText="Nome" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="Produto Entrada">
                                    <ItemTemplate>
                                        <asp:DropDownList ID="ddlXmlProdutoXDePara" runat="server" />
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="Unidade" HeaderText="Un." HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Lote" HeaderText="Lote" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Classificacao" HeaderText="Class." HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="DescricaoEmbalagem" HeaderText="Embalagem" HeaderStyle-HorizontalAlign="Left"
                                    ItemStyle-HorizontalAlign="Left">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="QtdeProgramada" DataFormatString="{0:N2}" HeaderText="Qtde Programada"
                                    HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="QtdeContratadoFiscal" DataFormatString="{0:N2}" HeaderText="Qtde Contratado"
                                    HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="QtdeEntregueFiscalAFixar" DataFormatString="{0:N2}" HeaderText="Qtde Entregue"
                                    HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="VlrNotaOficialAfixarBruto" DataFormatString="{0:N2}" HeaderText="Vlr Entregue R$"
                                    HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="VlrNotaMoedaAfixarBruto" DataFormatString="{0:N2}" HeaderText="Vlr Entregue $"
                                    HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="QtdeFixacao" DataFormatString="{0:N2}" HeaderText="Qtde Fixada"
                                    HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="VlrFixacaoOficial" DataFormatString="{0:N2}" HeaderText="Vlr Fixado R$"
                                    HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="SaldoQtdeAFixar" DataFormatString="{0:N2}" HeaderText="Saldo Qtde AFixar"
                                    HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                            </Columns>
                        </asp:GridView>
                    </div>
                </div>
            </asp:Panel>
            <div id="divFixacoes" runat="server">
                <div class="subtitulodiv">
                    Fixações Pendentes
                </div>
                <div class="row">
                    <div class="collbluc" style="width: 128px;">
                        Data Nota
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataNotaFix" runat="server" CssClass="calendario" Width="90px" />
                    </div>
                    <div class="collbluc" style="width: 128px;">
                        Data Saída/Entrada
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataMovimentoFix" runat="server" CssClass="calendario" Width="90px" />
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkNossaEmissaoFix" runat="server" AutoPostBack="True" Font-Bold="True"
                            OnCheckedChanged="chkNossaEmissaoFix_CheckedChanged" Text="Nossa Emissão" />
                    </div>
                    <div id="divNossaEmissao" runat="server">
                        <div class="collbluc" style="width: 128px;">
                            <asp:Label ID="lblNumNotaFix" runat="server" Text="Num. Nota" />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtNotaFix" CssClass="txtNumerico" runat="server" Width="80px" MaxLength="9" />
                        </div>
                        <div class="collbluc" style="width: 128px;">
                            <asp:Label ID="lblSerieFix" runat="server" Text="Série Nota" />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtSerieFix" runat="server" Width="50px" />
                        </div>
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc" style="width: 128px;">
                        Operação Comercial:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlOperacaoFixacao" runat="server" Width="600px" AutoPostBack="True" />
                    </div>
                </div>
                <div class="bordagrid" style="height: 160px;">
                    <asp:GridView ID="gridFixacao" runat="server" CellPadding="0" ForeColor="#333333"
                        GridLines="None" AutoGenerateColumns="False" Width="100%">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <EditRowStyle BackColor="#999999" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                        <Columns>
                            <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                            <asp:BoundField DataField="Nome" HeaderText="Cliente">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Pedido" HeaderText="Pedido">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="NomeMoeda" HeaderText="Moeda">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Movimento" DataFormatString="{0:dd-MM-yyyy}" HeaderText="Data">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Fixacao" HeaderText="Fixação">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Produto" HeaderText="Produto">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Operacao" HeaderText="Operacao">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Quantidade" HeaderText="Qtde">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Unitario" HeaderText="Unit.">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Valor" HeaderText="Vlr. Fixação" DataFormatString="{0:N2}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ValorNF" DataFormatString="{0:N2}" HeaderText="Vlr. Nota">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Saldo" DataFormatString="{0:N2}" HeaderText="Saldo">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <div id="divDevolucaoVNotas" runat="server">
                <div class="subtitulodiv">
                    Devolução várias Notas Fiscais
                </div>
                <div class="row">
                    <div class="collbluc" style="width: 128px;">
                        Operação:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlOperacaoDevNN" runat="server" Width="600px" AutoPostBack="True" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc" style="width: 128px;">
                        Sub-Operação:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlSubOperacaoDevNN" runat="server" Width="600px" AutoPostBack="True" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc" style="width: 128px;">
                        Data Nota
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataNotaDevNN" runat="server" CssClass="calendario" Width="90px" />
                    </div>
                    <div class="collbluc" style="width: 128px;">
                        Data Saída/Entrada
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataMovimentoDevNN" runat="server" CssClass="calendario" Width="90px" />
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkNossaEmissaoDevNN" runat="server" AutoPostBack="True" Font-Bold="True"
                            OnCheckedChanged="chkNossaEmissaoDevNN_CheckedChanged" Text="Nossa Emissão" />
                    </div>
                    <div id="divNossaEmissaoDevNN" runat="server">
                        <div class="collbluc" style="width: 128px;">
                            <asp:Label ID="lblNumNotaDevNN" runat="server" Text="Num. Nota" />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtNotaDevNN" CssClass="txtNumerico" runat="server" Width="80px" MaxLength="9" ToolTip="Número da Nota Fiscal." />
                        </div>
                        <div class="collbluc" style="width: 128px;">
                            <asp:Label ID="lblSerieDevNN" runat="server" Text="Série Nota" />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtSerieDevNN" runat="server" Width="50px" MaxLength="3" ToolTip="Série da Nota Fiscal." />
                        </div>
                    </div>
                </div>
                <div class="row">
                    <div class="row">
                        <div class="collbluc">
                            Grupo:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlGrupoProdutoDevNN" runat="server" AutoPostBack="True" Width="265px" />
                        </div>
                        <div class="coltxt">
                            <asp:LinkButton ID="lnkBuscaProdutoDevNN" runat="server" Height="20px" Width="20px">
                                <asp:Image ID="imgSearch" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                                    ToolTip="Consulta Produto" />
                            </asp:LinkButton>
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Produto:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlProdutosDevNN" runat="server" AutoPostBack="True" Width="285px" />
                        </div>
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Comercialização:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlUnidadeComercializacaoDevNN" runat="server" Width="150px" AutoPostBack="True" />
                    </div>
                    <div class="collbluc">
                        Quantidade
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtQuantidadeDevNN" runat="server" CssClass="txtDecimal4" Style="color: Blue; text-align: right" data-ToolTip="default" ToolTip="">0,0000</asp:TextBox>
                    </div>

                    <div class="collbluc">
                        Unitário
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtUnitarioDevNN" runat="server" CssClass="txtDecimal10" Style="color: Blue; text-align: right" TabIndex="5" data-ToolTip="default" ToolTip="">0,0000000000</asp:TextBox>
                    </div>

                    <div class="collbluc">
                        Valor
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtValorTotalDevNN" runat="server" CssClass="txtDecimal" Style="color: Blue; text-align: right"
                            Width="120px" />
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:Button ID="btnConfirmar" runat="server" CssClass="botao" Text="Confirmar" OnClick="btnConfirmar_Click"
                        Visible="False" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
