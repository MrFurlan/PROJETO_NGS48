<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AnaliseDeCreditoProdutorRural.aspx.vb" Inherits="NGS.Web.UI.AnaliseDeCreditoProdutorRural" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" language="javascript">
        function pageLoadSelecaoProduto() {
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
            pageLoadSelecaoProduto();
            var prmSelecaoProduto = Sys.WebForms.PageRequestManager.getInstance();
            prmSelecaoProduto.add_endRequest(pageLoadSelecaoProduto);
        });
        function Button2_onclick() {
        }

    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <table style="width: 100%; border: 0px none;">
                <tr>
                    <td class="titulotabela">Análise De Credito Produtor Rural
                    </td>
                </tr>
                <tr>
                    <td>
                        <ajaxToolkit:TabContainer ID="tcAnaliseCredito" runat="server" ActiveTabIndex="0"
                            Width="100%">
                            <ajaxToolkit:TabPanel ID="tbAnalise" runat="server">
                                <HeaderTemplate>
                                    Análise
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div style="overflow: auto; width: 100%;">
                                        <table style="width: 100%; border: 0px none;">
                                            <tr>
                                                <td colspan="4">
                                                    <div class="menu_acoes">
                                                        <div class="acoes">
                                                            <ul>
                                                                <li runat="server" class="iconNovo">
                                                                    <asp:LinkButton Text="Gravar" ID="lnkNovo" runat="server" />
                                                                </li>
                                                                <li runat="server" class="iconExcluir">
                                                                    <asp:LinkButton Text="Cancelar" ID="lnkCancelar" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                                                </li>
                                                                <li runat="server" class="iconNovo">
                                                                    <asp:LinkButton Text="Aprovar Simulação" ID="lnkAprovarSimulacao" runat="server" />
                                                                </li>
                                                                <li runat="server" class="iconNovo">
                                                                    <asp:LinkButton Text="Liberar Crédito" ID="lnkLiberarCredito" runat="server" />
                                                                </li>
                                                                <li runat="server" class="iconLimpar">
                                                                    <asp:LinkButton Text="Limpar" ID="lnkLimparAnalise" runat="server" />
                                                                </li>
                                                                <li runat="server" class="iconAjuda">
                                                                    <asp:LinkButton Text="Ajuda" ID="lnkAjudaB" runat="server" />
                                                                </li>
                                                            </ul>
                                                        </div>
                                                    </div>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="width: 100px;">
                                                    <div class="headerGray">
                                                        <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Ano:</span>
                                                        <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                    </div>
                                                </td>
                                                <td>
                                                    <asp:DropDownList ID="ddlAno" runat="server" AutoPostBack="True" EnableTheming="True"
                                                        OnSelectedIndexChanged="ddlAno_SelectedIndexChanged" Width="121px">
                                                        <asp:ListItem></asp:ListItem>
                                                        <asp:ListItem Value="2012"></asp:ListItem>
                                                        <asp:ListItem>2013</asp:ListItem>
                                                        <asp:ListItem>2014</asp:ListItem>
                                                        <asp:ListItem>2015</asp:ListItem>
                                                        <asp:ListItem>2016</asp:ListItem>
                                                        <asp:ListItem>2017</asp:ListItem>
                                                        <asp:ListItem>2018</asp:ListItem>
                                                        <asp:ListItem>2019</asp:ListItem>
                                                        <asp:ListItem>2020</asp:ListItem>
                                                    </asp:DropDownList>
                                                </td>
                                                <td style="width: 100px;">
                                                    <div class="headerGray">
                                                        <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Análise:</span>
                                                        <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                    </div>
                                                </td>
                                                <td>
                                                    <asp:Label ID="lblAnalise" runat="server" Font-Bold="True" Font-Size="12pt" ForeColor="Red" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <asp:Label ID="lblDefinicao" runat="server" Font-Bold="True" Font-Size="12pt" ForeColor="Red"
                                                        Width="400px" />
                                                </td>
                                                <td style="width: 100px;">
                                                    <div class="headerGray">
                                                        <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Situação:</span>
                                                        <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                    </div>
                                                </td>
                                                <td>
                                                    <asp:DropDownList ID="ddlSituacao" runat="server" Enabled="False" Width="213px">
                                                        <asp:ListItem Value="NORMAL">Normal</asp:ListItem>
                                                        <asp:ListItem Value="CANCELADA">Cancelada</asp:ListItem>
                                                        <asp:ListItem Value="APROVADA">Aprovada</asp:ListItem>
                                                        <asp:ListItem Value="LIBERADA">Liberada</asp:ListItem>
                                                    </asp:DropDownList>
                                                </td>
                                            </tr>
                                            <tr valign="bottom">
                                                <td class="titulotabela" colspan="2">
                                                    <asp:Button ID="btnIncluirCliente" runat="server" OnClick="btnIncluirCliente_Click"
                                                        Text=" + " />
                                                    Clientes
                                                </td>
                                                <td class="rotulo" rowspan="6" valign="top" colspan="2">
                                                    <table width="100%">
                                                        <tr>
                                                            <td class="titulotabela" colspan="2">Resumo Culturas
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo" style="background-color: #009933; color: white;">Receita Culturas
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtReceitaCulturas" runat="server" BackColor="#FFFF80" ReadOnly="True"
                                                                    Style="text-align: right;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo" style="background-color: #009933; color: white;">Outras Receitas
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtOutrasReceitas" runat="server" BackColor="#FFFF80" ReadOnly="True"
                                                                    Style="text-align: right;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo" style="background-color: #009933; color: white;">Total Receita
                                                            </td>
                                                            <td class="rotulo">
                                                                <asp:TextBox ID="txtTotalReceitas" runat="server" BackColor="#FFFF80" ReadOnly="True"
                                                                    Style="text-align: right;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo" style="background-color: #cc0000; color: white;">Custo Culturas
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtCustoCultura" runat="server" BackColor="#FFFF80" ReadOnly="True"
                                                                    Style="text-align: right;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo" style="background-color: #cc0000; color: white;">Custo Arrendamentos
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtCustoArrendamento" runat="server" BackColor="#FFFF80" ReadOnly="True"
                                                                    Style="text-align: right;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo" style="background-color: #cc0000; color: white;">Outros Custos
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtOutrosCustos" runat="server" BackColor="#FFFF80" ReadOnly="True"
                                                                    Style="text-align: right;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo" style="background-color: #cc0000; color: white;">Total Custos
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtTotalCustos" runat="server" BackColor="#FFFF80" ReadOnly="True"
                                                                    Style="text-align: right;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Saldo
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtSaldoAnalise" runat="server" BackColor="#FFFF80" ReadOnly="True"
                                                                    Style="text-align: right;" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <br />
                                                    <table width="100%">
                                                        <tr>
                                                            <td class="titulotabela" colspan="2">Rating
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Rating Credito %
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtPercRatingCredito" runat="server" BackColor="#FFFF80" ReadOnly="True"
                                                                    Width="40px" Style="text-align: right;" /><asp:Label ID="lblRatingCredito"
                                                                        runat="server" BackColor="Black" Font-Bold="True" Font-Names="Perpetua Titling MT"
                                                                        Font-Size="Medium" ForeColor="White" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Coef. Red. Risco Cultura %
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtCoefRedutor" runat="server" Width="40px" BackColor="#FFFF80"
                                                                    ReadOnly="True" Style="text-align: right;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Potencial Compra Portifolio
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtPotencialCompraPortifolio" runat="server" BackColor="#FFFF80"
                                                                    Style="text-align: right;" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td></td>
                                                            <td></td>
                                                        </tr>
                                                        <tr>
                                                            <td class="rotulo">Credito Concedido
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtCreditoConcedido" runat="server" BackColor="#FFFF80" Style="text-align: right;" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <br />
                                                    <table border="0" class="borda" width="240">
                                                        <tr>
                                                            <td align="left">Aprovação<br />
                                                                <asp:Image ID="imgUsuarioIncl" runat="server" Height="20px" ImageAlign="AbsMiddle"
                                                                    ImageUrl="~/images/man2.png" Width="18px" />
                                                                <asp:Label ID="lblUsuarioApr" runat="server" Font-Bold="False" Width="190px" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left">Liberação<br />
                                                                <asp:Image ID="imgUsuarioAlt" runat="server" Height="20px" ImageAlign="AbsMiddle"
                                                                    ImageUrl="~/images/man2.png" Width="18px" />
                                                                <asp:Label ID="lblUsuarioLib" runat="server" Font-Bold="False" Width="190px" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td align="left" style="height: 22px">Cancelamento<br />
                                                                <asp:Image ID="Image1" runat="server" Height="20px" ImageAlign="AbsMiddle" ImageUrl="~/images/man2.png"
                                                                    Width="18px" />
                                                                <asp:Label ID="lblUsuarioCan" runat="server" Font-Bold="False" Width="190px" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr valign="top">
                                                <td colspan="2">
                                                    <asp:GridView ID="gridClientes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                        ForeColor="#333333" GridLines="None">
                                                        <AlternatingRowStyle BackColor="White" />
                                                        <Columns>
                                                            <asp:BoundField DataField="CodigoCliente" HeaderText="Cliente" />
                                                            <asp:BoundField DataField="NomeCliente" HeaderText="Nome" />
                                                            <asp:TemplateField ShowHeader="False">
                                                                <ItemTemplate>
                                                                    <asp:ImageButton ID="ImgExcluir" runat="server" CausesValidation="False" CommandName="Delete"
                                                                        ImageUrl="~/images/Cancelar.jpg" OnClick="ImgExcluir_Click" Text="Excluir" />
                                                                </ItemTemplate>
                                                            </asp:TemplateField>
                                                        </Columns>
                                                        <EditRowStyle BackColor="#2461BF" />
                                                        <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                                        <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                                        <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                                        <RowStyle BackColor="#EFF3FB" />
                                                        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                                                    </asp:GridView>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="titulotabela" colspan="2">Culturas
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <div style="overflow: auto;">
                                                        <asp:GridView ID="gridCulturas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                            ForeColor="#333333" GridLines="None">
                                                            <AlternatingRowStyle BackColor="White" />
                                                            <Columns>
                                                                <asp:BoundField DataField="CodigoSafra" HeaderText="Safra" />
                                                                <asp:BoundField DataField="NomeCultura" HeaderText="Cultura" />
                                                                <asp:BoundField DataField="AreaPlantio" DataFormatString="{0:N2}" HeaderText="Área de Plantio (ha)">
                                                                    <HeaderStyle HorizontalAlign="Right" />
                                                                    <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="PercCultura" DataFormatString="{0:N2}" HeaderText="% Cult.">
                                                                    <HeaderStyle HorizontalAlign="Right" />
                                                                    <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                                </asp:BoundField>
                                                                <asp:TemplateField HeaderText="Risco Cultura">
                                                                    <ItemTemplate>
                                                                        <asp:DropDownList ID="ddlRiscoCultura" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlRiscoCultura_SelectedIndexChanged"
                                                                            SelectedValue='<%# eval("RiscoCultura") %>'>
                                                                            <asp:ListItem Value="1">Baixo</asp:ListItem>
                                                                            <asp:ListItem Value="2">Medio</asp:ListItem>
                                                                            <asp:ListItem Value="3">Alto</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </ItemTemplate>
                                                                </asp:TemplateField>
                                                                <asp:BoundField DataField="Producao" DataFormatString="{0:N2}" HeaderText="Produção">
                                                                    <HeaderStyle HorizontalAlign="Right" />
                                                                    <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="ReceitaCultura" DataFormatString="{0:N2}" HeaderText="Receita">
                                                                    <HeaderStyle HorizontalAlign="Right" />
                                                                    <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="CustoPortifolio" DataFormatString="{0:N2}" HeaderText="Custo de Portifólio">
                                                                    <HeaderStyle HorizontalAlign="Right" />
                                                                    <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="CustoCultura" DataFormatString="{0:N2}" HeaderText="Custo de Produção">
                                                                    <HeaderStyle HorizontalAlign="Right" />
                                                                    <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                                </asp:BoundField>
                                                            </Columns>
                                                            <EditRowStyle BackColor="#2461BF" />
                                                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                                            <RowStyle BackColor="#EFF3FB" />
                                                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                                                        </asp:GridView>
                                                    </div>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="titulotabela" colspan="2">Perguntas
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <div style="overflow: auto;">
                                                        <asp:GridView ID="gridPerguntas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                            ForeColor="#333333" GridLines="None">
                                                            <AlternatingRowStyle BackColor="White" />
                                                            <Columns>
                                                                <asp:BoundField DataField="DescPergunta" HeaderText="Pergunta" />
                                                                <asp:BoundField DataField="PercPeso" HeaderText="Peso (%)">
                                                                    <HeaderStyle HorizontalAlign="Right" />
                                                                    <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                                </asp:BoundField>
                                                                <asp:TemplateField>
                                                                    <ItemTemplate>
                                                                        <asp:DropDownList ID="ddlResposta" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlResposta_SelectedIndexChanged"
                                                                            SelectedValue='<%# eval("Resposta") %>'>
                                                                            <asp:ListItem Value="3">Bom</asp:ListItem>
                                                                            <asp:ListItem Value="2">Regular</asp:ListItem>
                                                                            <asp:ListItem Value="1">Ruim</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </ItemTemplate>
                                                                </asp:TemplateField>
                                                            </Columns>
                                                            <EditRowStyle BackColor="#2461BF" />
                                                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                                                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                                                            <RowStyle BackColor="#EFF3FB" />
                                                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                                                        </asp:GridView>
                                                    </div>
                                                </td>
                                            </tr>
                                        </table>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="tbConsulta" runat="server">
                                <HeaderTemplate>
                                    Consulta
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <table style="width: 100%; border: 0px none;">
                                        <tr>
                                            <td class="titulotabela">Parâmetros da Consulta
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div class="menu_acoes">
                                                    <div class="acoes">
                                                        <ul>
                                                            <li class="iconRelatorio" runat="server">
                                                                <asp:LinkButton Text="Relatório" ID="lnkRelatorio" runat="server" />
                                                            </li>
                                                            <li class="iconConsultar" runat="server">
                                                                <asp:LinkButton Text="Consultar" ID="lnkConsultar" runat="server" />
                                                            </li>
                                                            <li class="iconLimpar" runat="server">
                                                                <asp:LinkButton Text="Limpar" ID="lnkLimpar" runat="server" />
                                                            </li>
                                                            <li runat="server">
                                                                <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                                                            </li>
                                                        </ul>
                                                    </div>
                                                </div>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div id="divConsulta" class="accordion">
                                                    <h3>Consulta
                                                    </h3>
                                                    <table>
                                                        <tr>
                                                            <td style="width: 70px; vertical-align: top;">
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Análise:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtAnaliseConsulta" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Ano:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:DropDownList ID="ddlAnoConsulta" runat="server" AutoPostBack="True" EnableTheming="True"
                                                                    Width="121px">
                                                                    <asp:ListItem></asp:ListItem>
                                                                    <asp:ListItem Value="2012"></asp:ListItem>
                                                                    <asp:ListItem>2013</asp:ListItem>
                                                                    <asp:ListItem>2014</asp:ListItem>
                                                                    <asp:ListItem>2015</asp:ListItem>
                                                                    <asp:ListItem>2016</asp:ListItem>
                                                                    <asp:ListItem>2017</asp:ListItem>
                                                                    <asp:ListItem>2018</asp:ListItem>
                                                                    <asp:ListItem>2019</asp:ListItem>
                                                                    <asp:ListItem>2020</asp:ListItem>
                                                                </asp:DropDownList>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Cliente:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtClienteConsulta" runat="server" Enabled="False" Width="485px" />
                                                                <asp:Button ID="btnCliente" runat="server" OnClick="btnCliente_Click" Text=" &gt; " />
                                                                <asp:HiddenField ID="HCliente" runat="server" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Situação:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:CheckBox ID="chkNormal" runat="server" Text="Normal" />
                                                                <asp:CheckBox ID="chkCancelada" runat="server" Text="Cancelada" />
                                                                <asp:CheckBox ID="chkAprovada" runat="server" Text="Aprovada" />
                                                                <asp:CheckBox ID="chkLiberada" runat="server" Text="Liberada" />
                                                                &nbsp;
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Aprovação:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtDataAprDe" runat="server" CssClass="calendario" />
                                                                a
                                                                <asp:TextBox ID="txtDataAprAte" runat="server" CssClass="calendario" />
                                                                <asp:CheckBox ID="chkUsarPeriodoAprovacao" runat="server" Text="Usar Período" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Liberação:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtDataLibDe" runat="server" CssClass="calendario" />
                                                                a
                                                                <asp:TextBox ID="txtDataLibAte" runat="server" CssClass="calendario" />
                                                                <asp:CheckBox ID="chkUsarPeriodoLiberacao" runat="server" Text="Usar Período" />
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Usuário:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtUsuario" runat="server" />
                                                                &nbsp;
                                                                <asp:CheckBox ID="chkUsuarioAprovacao" runat="server" Text="Aprovacao" />
                                                                <asp:CheckBox ID="chkUsuarioLiberacao" runat="server" Text="Liberação" />
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </div>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div style="min-width: 100%; max-width: 900px; width: 150%; overflow: auto; height: 450px;">
                                                    <asp:GridView ID="gridConsulta" runat="server" CellPadding="4" ForeColor="#333333"
                                                        GridLines="None" CssClass="bordasimples" OnSelectedIndexChanged="gridConsulta_SelectedIndexChanged"
                                                        AutoGenerateColumns="False" Width="100%">
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <EditRowStyle BackColor="#999999" />
                                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                        <Columns>
                                                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                                            <asp:TemplateField>
                                                                <ItemTemplate>
                                                                    <asp:ImageButton ID="imgImpressaoPend" runat="server" Height="20px" ImageUrl="~/images/impressora.JPG"
                                                                        OnClick="imgImpressaoPend_Click" />
                                                                </ItemTemplate>
                                                            </asp:TemplateField>
                                                            <asp:BoundField DataField="Analise_Id" HeaderText="ID">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="Ano_Id" HeaderText="Ano" />
                                                            <asp:BoundField DataField="DefinicaoAno_Id" HeaderText="Definição" />
                                                            <asp:BoundField DataField="Situacao" HeaderText="Situação" />
                                                            <asp:BoundField DataField="Nome" HeaderText="Nome" />
                                                            <asp:BoundField DataField="Plantios" HeaderText="Nº de Plantios">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="AreaPlantio" DataFormatString="{0:N2}" HeaderText="Área de Plantio">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="AreaArrendada" DataFormatString="{0:N2}" HeaderText="Área de Arrend.">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="CustoArrendamento" DataFormatString="{0:N2}" HeaderText="Custo Arrend.">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="CustoPortifolio" DataFormatString="{0:N2}" HeaderText="Custo Portif.">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="CustoCultura" DataFormatString="{0:N2}" HeaderText="Custo Cult.">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="OutrasDespesas" DataFormatString="{0:N2}" HeaderText="Outras Despesas">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="OutrasReceitas" DataFormatString="{0:N2}" HeaderText="Outras Receitas">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="ReceitaCultura" DataFormatString="{0:N2}" HeaderText="Receita Cult.">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="LimiteCredito" HeaderText="Lim. Cred.">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PercLimCredito" DataFormatString="{0:N2}" HeaderText="(%) Cred.">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="PercRedutorRiscoCultura" DataFormatString="{0:N2}" HeaderText="(%) Risco Cult.">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                            <asp:BoundField DataField="CreditoAnalise" DataFormatString="{0:N2}" HeaderText="Crédito Ánalise">
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                            </asp:BoundField>
                                                        </Columns>
                                                    </asp:GridView>
                                                </div>
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
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
