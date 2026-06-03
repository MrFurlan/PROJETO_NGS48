<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AlterarNotaFiscal.aspx.vb" Inherits="NGS.Web.UI.AlterarNotaFiscal" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <br />
    <asp:ScriptManager ID="scrmngAlterarNotaFiscal" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAlterarNotaFiscal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Alterar Nota Fiscal
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
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
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="711px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="711px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="672px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=" > " UseSubmitBehavior="False" OnClick="btnCliente_Click"
                        CssClass="btn" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="96px"
                        data-ToolTip="default" ToolTip="Periodo Inicial" />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="96px"
                        CausesValidation="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtES" runat="server" MaxLength="3" Width="96px" />
                </div>
                <div class="collbl" style="margin-left: 21px;">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNota" CssClass="numerico" runat="server" Width="96px" />
                </div>
                <div class="collbl">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerie" runat="server" MaxLength="3" Width="96px" />
                </div>
            </div>
            <div id="pnlDados" runat="server">
                <div class="painelleft" style="width: 45.5%; margin-right: 4px;">
                    <div class="row">
                        <div class="collbl">
                            Pedido:
                        </div>
                        <div class="coltxt">
                            <asp:Button ID="btnPedido" runat="server" Text=" > " UseSubmitBehavior="False" Enabled="False" />
                            <asp:Label ID="txtPedido" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                            <asp:ImageButton ID="imgExtratoPedido" runat="server" Width="23px" Height="23px"
                                ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ImageAlign="AbsMiddle"
                                data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido" Enabled="False" OnClick="imgExtratoPedido_Click"></asp:ImageButton>
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">Romaneio:</div>
                        <div class="coltxt">
                            <asp:Button ID="btnRomaneio" OnClick="btnRomaneio_Click" runat="server" Text=" > "
                                UseSubmitBehavior="False"></asp:Button>
                            <asp:Label ID="txtRomaneio" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Pesagem:
                        </div>
                        <div class="coltxt">
                            <asp:Label ID="txtPesagem" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Finalidade:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="cmbFinalidade" runat="server"
                                Width="300px" AutoPostBack="True" OnSelectedIndexChanged="cmbFinalidade_SelectedIndexChanged" />
                        </div>
                    </div>
                     <div class="row">
                        <div class="collbl">
                            Vencimento:
                        </div>
                        <div class="coltxt">
                            <asp:Button ID="BtnVencimentos" OnClick="BtnVencimentos_Click" runat="server" Text="GERAR FINANCEIRO"
                                UseSubmitBehavior="False"></asp:Button>
                            <asp:Label ID="txtTitulo" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                        </div>
                    </div>
                </div>
                <div class="painelleft" style="width: 53.5%;">
                    <div class="subtitulodiv">
                        Resumo Nota / Alteração
                    </div>
                    <div class="painelleft" style="width: 49.5%;">
                        <div class="row">
                            <div class="collbl">
                                Total dos Produtos:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblTotalProduto" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Total da Nota Fiscal:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblTotalNota" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Quantidade Física:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblFisico" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                            </div>
                        </div>

                    </div>
                    <div class="painelleft" style="width: 49.5%;">
                        <div class="row">
                            <div class="collbl">
                                Novo Total Prd.:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblNovoTotalProduto" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Novo Total NF:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblNovoTotalNota" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Nova Qtde Física:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblNovoFisico" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <asp:Panel ID="pnlItem" runat="server"  Width="100%">
                    <div class="bordagrid" style="height: 150px">
                        <asp:GridView ID="gridItem" runat="server" AutoGenerateColumns="False" CellPadding="4"
                            ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridItem_SelectedIndexChanged"
                            OnRowCommand="gridItem_RowCommand" OnRowCreated="gridItem_RowCreated">
                            <EditRowStyle BackColor="#999999" />
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <Columns>
                                <asp:CommandField ButtonType="Button" InsertText="" SelectText=" &gt; " ShowSelectButton="True" />
                                <asp:BoundField DataField="Produto" HeaderText="Codigo">
                                    <HeaderStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="NomeProduto" HeaderText="Produto">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle Width="200px" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Lote" HeaderText="Lote" />
                                <asp:BoundField DataField="Classificacao" HeaderText="Class." />
                                <asp:BoundField DataField="Embalagem" HeaderText="Emb." />
                                <asp:BoundField DataField="Saldo" HeaderText="Saldo">
                                    <HeaderStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="Qtde Física">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtQuantidadeFItem" runat="server" CssClass="txtDecimal4" Text='<%# Eval("QuantidadeF", "{0:N4}") %>'
                                            Enabled="False" Width="85px" />
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Right" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Qtde Fiscal">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtQuantidadeItem" runat="server" CssClass="txtDecimal4" Text='<%# Eval("Quantidade", "{0:N4}") %>' Width="85px" />
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Right" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Unitário">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtUnitarioItem" CssClass="txtDecimal10" runat="server" Text='<%# Eval("Unitario", "{0:N10}") %>' Width="85px" />
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Right" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Total">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtTotalItem" runat="server" Enabled="False" CssClass="txtDecimal"
                                            Text='<%# Eval("Total", "{0:N2}") %>' Width="85px" />
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Right" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Peças/Meios">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtNumeroPecas" runat="server" Enabled="False" Text='<%# Eval("NumeroPecas", "{0:N0}") %>' Width="50px" />
                                    </ItemTemplate>
                                    <HeaderStyle HorizontalAlign="Right" />
                                </asp:TemplateField>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:Button ID="btnItem" runat="server" Text="OK" CommandName="OK" UseSubmitBehavior="False" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                    </div>
                </asp:Panel>
          
            <asp:Panel ID="pnlGeral" runat="server" Width="100%" ScrollBars="Auto">
                <table style="border: 0; width: 100%">
                    <tr>
                        <td>
                            <div class="bordagrid" style="height: 250px">
                                <table style="width: 100%;">
                                <tr>
                                    <td>
                                        <strong>OPERAÇÃO</strong></span></label>&nbsp;<br />
                                        <asp:DropDownList ID="cmbSubOperacao" runat="server" Enabled="False"
                                            Width="500px">
                                        </asp:DropDownList>
                                        &nbsp;<asp:Button ID="btnFisico" runat="server" UseSubmitBehavior="False" Text="Físico"
                                            BackColor="Green" BorderStyle="None" Style="cursor: pointer; font-family: Tahoma,Arial,Helvetica,sans-serif; font-size: 11px; height: 24px; text-align: center; width: 96px;"
                                            Font-Bold="True"
                                            ForeColor="White" Visible="False" />&nbsp;
                                        <asp:Button ID="btnFiscal" runat="server" Text="Fiscal" UseSubmitBehavior="False"
                                            BackColor="Green" BorderStyle="None" Style="cursor: pointer; font-family: Tahoma,Arial,Helvetica,sans-serif; font-size: 11px; height: 24px; text-align: center; width: 96px;"
                                            Font-Bold="True"
                                            ForeColor="White" Visible="False" />
                                        <asp:Button ID="btnAtualizar" runat="server" OnClick="btnAtualizar_Click" Text="Atualizar"
                                            CssClass="botao"
                                            UseSubmitBehavior="False" />
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <table border="1" style="width: 100%;">
                                            <tr>
                                                <td>
                                                    <strong>NATUREZA DA OPERAÇÃO</strong>
                                                </td>
                                                <td>
                                                    <strong>SITUAÇÃO TRIBUTÁRIA ICMS</strong>
                                                </td>
                                                <td>
                                                    <strong>SITUAÇÃO TRIBUTÁRIA IPI</strong>
                                                </td>
                                                <td>
                                                    <strong>SITUAÇÃO TRIBUTÁRIA PIS/COFINS</strong>
                                                </td>
                                                <td>
                                                    <strong>OperaçãoXEstado</strong>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:Label ID="txtNaturezaDaOperacao" runat="server" ForeColor="Blue" Font-Bold="False"
                                                        BorderStyle="None" />
                                                </td>
                                                <td>
                                                    <asp:Label ID="txtSituacaoTributaria" runat="server" ForeColor="Blue" Font-Bold="False"
                                                        BorderStyle="None" />
                                                </td>
                                                <td>
                                                    <asp:Label ID="txtSituacaoTributariaIPI" runat="server" ForeColor="Blue" Font-Bold="False"
                                                        BorderStyle="None" />
                                                </td>
                                                <td>
                                                    <asp:Label ID="txtSituacaoTributariaPISCOFINS" runat="server" ForeColor="Blue" Font-Bold="False"
                                                        BorderStyle="None" />
                                                </td>
                                                <td>
                                                    <asp:Label ID="txtOperacaoXEstado" runat="server" ForeColor="Blue" Font-Bold="False"
                                                        BorderStyle="None" />
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Panel ID="pnlEncargos" runat="server" Width="100%" ScrollBars="Auto" CssClass="bordasimples">
                                            <asp:GridView ID="gridEncargos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                ForeColor="#333333" GridLines="None" Width="100%">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:BoundField DataField="Codigo" HeaderText="Encargo">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CodigoOperacao" HeaderText="Op">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CodigoSubOperacao" HeaderText="So">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CodigoGrupoProduto" HeaderText="Grupo">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CodigoProduto" HeaderText="Produto">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="EstadoOrigem" HeaderText="Origem">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="EstadoDestino" HeaderText="Destino">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="SituacaoTributaria" HeaderText="Icms">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="SituacaoTributariaIPI" HeaderText="Ipi">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="SituacaoTributariaPISCOFINS" HeaderText="Pis/Cofins">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:TemplateField HeaderText="Base">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtBaseEncargoItem" runat="server" BorderStyle="None" Enabled="False"
                                                                CssClass="txtDecimal" Text='<%# Eval("Base", "{0:N2}") %>' Width="100px" />
                                                        </ItemTemplate>
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="Percentual" DataFormatString="{0:N9}" HeaderText="% Calc.">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="PercentualExibicao" DataFormatString="{0:N9}" HeaderText="% Exib.">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:TemplateField HeaderText="Valor">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtValorEncargoItem" runat="server" BorderStyle="None" Enabled="False"
                                                                CssClass="txtDecimal" Text='<%# Eval("Valor", "{0:N2}") %>' Width="100px" />
                                                        </ItemTemplate>
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="Sinal" HeaderText="">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                </Columns>
                                                <EditRowStyle BackColor="#999999" />
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                            </asp:GridView>
                                        </asp:Panel>
                                    </td>
                                </tr>
                            </table>
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <div class="bordagrid" style="height: 350px">
                                <table width="100%">
                                <tr>
                                    <td>
                                        <strong>NOVA OPERAÇÃO</strong></span></label>&nbsp;<br />
                                        <asp:DropDownList ID="cmbNovaSubOperacao" runat="server" AutoPostBack="True" Enabled="False"
                                            Width="500px" OnSelectedIndexChanged="cmbNovaSubOperacao_SelectedIndexChanged">
                                        </asp:DropDownList>
                                        &nbsp;<asp:Button ID="btnFisico2" runat="server" Text="Físico" BackColor="Green"
                                            BorderStyle="None" Style="cursor: pointer; font-family: Tahoma,Arial,Helvetica,sans-serif; font-size: 11px; height: 24px; text-align: center; width: 96px;"
                                            Font-Bold="True"
                                            ForeColor="White" UseSubmitBehavior="False" Visible="False" />&nbsp;<asp:Button ID="btnFiscal2"
                                                runat="server" BackColor="Green" BorderStyle="None" Style="cursor: pointer; font-family: Tahoma,Arial,Helvetica,sans-serif; font-size: 11px; height: 24px; text-align: center; width: 96px;"
                                                Font-Bold="True"
                                                ForeColor="White" Text="Fiscal" UseSubmitBehavior="False" Visible="False" />
                                    </td>
                                </tr>
                                <tr id="panelNovaOperacao" runat="server" visible="false">
                                    <td>
                                        <table border="1" style="width: 100%;">
                                            <tr>
                                                <td>
                                                    <strong>NATUREZA DA OPERAÇÃO</strong>
                                                </td>
                                                <td>
                                                    <strong>SITUAÇÃO TRIBUTÁRIA ICMS</strong>
                                                </td>
                                                <td>
                                                    <strong>SITUAÇÃO TRIBUTÁRIA IPI</strong>
                                                </td>
                                                <td>
                                                    <strong>SITUAÇÃO TRIBUTÁRIA PIS/COFINS</strong>
                                                </td>
                                                <td>
                                                    <strong>OperaçãoXEstado</strong>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:Label ID="txtNovaNaturezaDaOperacao" runat="server" ForeColor="Blue" Font-Bold="False"
                                                        BorderStyle="None" />
                                                </td>
                                                <td>
                                                    <asp:Label ID="txtNovaSituacaoTributaria" runat="server" ForeColor="Blue" Font-Bold="False"
                                                        BorderStyle="None" />
                                                </td>
                                                <td>
                                                    <asp:Label ID="txtNovaSituacaoTributariaIPI" runat="server" ForeColor="Blue" Font-Bold="False"
                                                        BorderStyle="None" />
                                                </td>
                                                <td>
                                                    <asp:Label ID="txtNovaSituacaoTributariaPISCOFINS" runat="server" ForeColor="Blue" Font-Bold="False"
                                                        BorderStyle="None" />
                                                </td>
                                                <td>
                                                    <asp:Label ID="txtNovaOperacaoXEstado" runat="server" ForeColor="Blue" Font-Bold="False"
                                                        BorderStyle="None" />
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Panel ID="pnlEncargosNovos" runat="server" Width="100%" ScrollBars="Auto">
                                            <asp:GridView ID="gridEncargosNovos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                ForeColor="#333333" GridLines="None" Width="100%">
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <Columns>
                                                    <asp:BoundField DataField="Codigo" HeaderText="Encargo">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CodigoOperacao" HeaderText="Op">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CodigoSubOperacao" HeaderText="So">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CodigoGrupoProduto" HeaderText="Grupo">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CodigoProduto" HeaderText="Produto">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="EstadoOrigem" HeaderText="Origem">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="EstadoDestino" HeaderText="Destino">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="SituacaoTributaria" HeaderText="Icms">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="SituacaoTributariaIPI" HeaderText="Ipi">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="SituacaoTributariaPISCOFINS" HeaderText="Pis/Cofins">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:TemplateField HeaderText="Base">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtBaseEncargoItem" runat="server" BorderStyle="None" Enabled="False"
                                                                CssClass="txtDecimal" Text='<%# Eval("Base", "{0:N2}") %>' Width="100px" />
                                                        </ItemTemplate>
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="Percentual" DataFormatString="{0:N9}" HeaderText="% Calc.">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="PercentualExibicao" DataFormatString="{0:N9}" HeaderText="% Exib.">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:TemplateField HeaderText="Valor">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtValorEncargoItem" runat="server" BorderStyle="None" Enabled="False"
                                                                CssClass="txtDecimal" Text='<%# Eval("Valor", "{0:N2}") %>' Width="100px" />
                                                        </ItemTemplate>
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="Sinal" HeaderText="">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                </Columns>
                                                <EditRowStyle BackColor="#999999" />
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                            </asp:GridView>
                                        </asp:Panel>
                                    </td>
                                </tr>                                
                            </table>
                            </div>
                        </td>
                    </tr>
                    
                </table>
            </asp:Panel>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidosXNotas ID="ucConsultaPedidosXNotas" runat="server" />
    <uc:NotaFiscalXClassificacao ID="ucNotaFiscalXClassificacao" runat="server" />
    <uc:ConsultaRomaneios ID="ucConsultaRomaneios" runat="server" />
    <uc:Vencimentos ID="ucVencimentos" runat="server" />
</asp:Content>
