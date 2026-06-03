<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RegistrosDeICMS.aspx.vb" Inherits="NGS.Web.UI.RegistrosDeICMS" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collblwidth {
            width: 130px;
        }
    </style>
    <script type="text/javascript">

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
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmgRegistrosDeICMS" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRegistrosDeICMS" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Registros Fiscais Icms
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server"
                                OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged1"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="634px" OnSelectedIndexChanged="DdlEmpresa_SelectedIndexChanged1"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Processo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlProcesso" runat="server" Width="634px" OnSelectedIndexChanged="DdlProcesso_SelectedIndexChanged1"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt" style="width: 317px;">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="96px" data-ToolTip="default" ToolTip="Informar de acordo com o período de apuração." />
                    &nbsp;à&nbsp;
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="96px" data-ToolTip="default" ToolTip="Informar de acordo com o período de apuração." />
                </div>
                <div class="collbl collblwidth">
                    Livro:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtLivro" runat="server" Width="96px" CssClass="txtNumerico" data-ToolTip="default" ToolTip="Informar o número do Livro." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Código Receita:
                </div>
                <div class="coltxt" style="width: 317px;">
                    <asp:TextBox ID="txtCodigoDaReceita" runat="server" CssClass="txtNumerico" data-ToolTip="default" ToolTip="Informar o código da receita." />
                </div>
                <div class="collbl collblwidth">
                    Folha:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFolha" runat="server" Width="96px" CssClass="txtNumerico" data-ToolTip="default" ToolTip="Informar o número da folha." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Parâmetros:
                </div>
                <div class="coltxt" style="line-height: 15px; margin-bottom: 4px;">
                    <asp:CheckBox ID="CkExcell" runat="server" Text="Gerar Excel" data-ToolTip="default" ToolTip="Marcar as opções para parametrizar o registro e os relatórios de conferência." />
                    <br />
                    <asp:CheckBox ID="ckConferenciaCfopIcms" runat="server" Text="Conferência Relatórios 6 e 7 Entradas/Saídas Por UF/CFOP" data-ToolTip="default" ToolTip="Marcar as opções para parametrizar o registro e os relatórios de conferência." />
                    <br />
                    <asp:CheckBox ID="CkConsNotasCompServicos" runat="server" BorderStyle="None" Checked="True"
                        Text="Considerar Notas Compostas Serviço" data-ToolTip="default" ToolTip="Marcar as opções para parametrizar o registro e os relatórios de conferência." />
                </div>
                <div class="collbl collblwidth">
                    Vencto da Obrigação:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtVenctoDaObrigacao" runat="server" CssClass="calendario" Width="96px" data-ToolTip="default" ToolTip="Informar a data de vencimento da obrigação." />
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
                    Folha inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFolhaInicial" runat="server" CssClass="texto" MaxLength="3" Style="text-transform: uppercase"
                        TabIndex="8" Width="36px" data-ToolTip="default" ToolTip="Informar a folha inicial" />
                </div>
                <div class="collbl">
                    Folha final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFolhaFinal" runat="server" CssClass="texto" MaxLength="3" Style="text-transform: uppercase"
                        TabIndex="8" Width="36px" data-ToolTip="default" ToolTip="Informar a folha final" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="hdfCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtNomeCliente" runat="server" ReadOnly="True" TabIndex="3" Width="300px"
                        Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultaClientes" runat="server" TabIndex="4" Text="&gt;" CssClass="btn"
                        UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
                <div class="collbl colw">
                    Número/Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumeroNota" runat="server" TabIndex="7" CssClass="txtNumerico9"
                        Width="88px" data-ToolTip="default" ToolTip="Informar o número da nota." />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerie" runat="server" CssClass="texto" MaxLength="3" Style="text-transform: uppercase"
                        TabIndex="8" Width="36px" data-ToolTip="default" ToolTip="Informar a série da nota." />
                </div>
            </div>
            <div class="bordagrid" style="height: 400px;">
                <asp:GridView ID="GridOpcoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridOpcoes_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                        <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="bordagrid" style="height: 105px;" runat="server" visible="false">
                <asp:GridView ID="GridTermos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridTermos_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                            <ItemStyle Width="100px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
            <table style="width: 100%; border: 0px none;">
                <tr>
                    <td colspan="4">
                        <div id="Div1" style="max-width: 100%;">
                        </div>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:RegistrosIcmsAjustaResumo ID="ucRegistrosIcmsAjustaResumo" runat="server" />
    <uc:DarDiferencialDeAliquota ID="ucDarDiferencialDeAliquota" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
