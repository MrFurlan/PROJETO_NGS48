<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CRM.aspx.vb" Inherits="NGS.Web.UI.CRM" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:Panel ID="Panel1" runat="server" Height="100%" Width="100%">
                <table style="width: 100%; border: 0px none;">
                    <tr>
                        <td colspan="3" style="font-weight: bold; font-size: 32px; color: white; font-family: Arial;
                            background-color: #003366; text-align: center">
                            Critérios de Segmentação dos Clientes
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <ajaxToolkit:TabContainer ID="TabCRM" runat="server" ActiveTabIndex="1" Width="100%">
                                <ajaxToolkit:TabPanel ID="TabPanel4" runat="server" HeaderText="TabPanel4">
                                    <HeaderTemplate>
                                        Consulta
                                    </HeaderTemplate>
                                    <ContentTemplate>
                                        <asp:Panel ID="Panel8" runat="server" Height="600px" Width="1024px">
                                            <table>
                                                <tr>
                                                    <td colspan="3" style="font-weight: bold; font-size: 22px; color: black; font-family: Arial;
                                                        background-color: #C0C0C0;">
                                                        Consulta
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="2">
                                                        Ano:
                                                        <asp:DropDownList ID="ddlAnoConsulta" runat="server" AutoPostBack="True" Height="19px"
                                                            Width="94px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                        <asp:CheckBox ID="chkImprimirCrm" runat="server" Checked="True" Text="Imprimir somente clientes CRM" />
                                                        &nbsp;<asp:Button ID="btnNovaAnalise" runat="server" CssClass="botao" Text="Nova Analise" />
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="3">
                                                        <asp:GridView ID="gridConsulta" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                            ForeColor="#333333" GridLines="None">
                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                            <Columns>
                                                                <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                                                <asp:BoundField DataField="CodigoEmpresa" HeaderText="Codigo" />
                                                                <asp:BoundField DataField="EndEmpresa" HeaderText="End" />
                                                                <asp:BoundField DataField="NomeEmpresa" HeaderText="Nome" />
                                                                <asp:CheckBoxField DataField="Consolidado" HeaderText="Consolidado">
                                                                    <ItemStyle HorizontalAlign="Center" />
                                                                </asp:CheckBoxField>
                                                                <asp:TemplateField HeaderText="Imprimir">
                                                                    <ItemTemplate>
                                                                        <asp:ImageButton ID="imgImpressaoPend" runat="server" Height="20px" ImageUrl="~/images/impressora.JPG"
                                                                            OnClick="imgImpressaoPend_Click" />
                                                                    </ItemTemplate>
                                                                    <ItemStyle HorizontalAlign="Center" />
                                                                </asp:TemplateField>
                                                                <asp:TemplateField HeaderText="Excluir">
                                                                    <ItemTemplate>
                                                                        <asp:ImageButton ID="imgExcluir" runat="server" Height="20px" ImageUrl="~/Images/desmarcar.ico"
                                                                            OnClick="imgExcluir_Click" />
                                                                    </ItemTemplate>
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
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        &#160;&#160;
                                                    </td>
                                                    <td>
                                                        &#160;&#160;
                                                    </td>
                                                    <td>
                                                        &#160;&#160;
                                                    </td>
                                                </tr>
                                            </table>
                                        </asp:Panel>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                                    <HeaderTemplate>
                                        Parâmetros da Curva
                                    </HeaderTemplate>
                                    <ContentTemplate>
                                        <table style="width: 100%; border: 0px none;">
                                            <tr>
                                                <td style="width: 95%">
                                                    <div id="divCurvaABC" class="accordion">
                                                        <h3>
                                                            Curva ABC
                                                        </h3>
                                                        <table width="100%">
                                                            <tr>
                                                                <td class="primario">
                                                                    Empresa Venda
                                                                </td>
                                                                <td>
                                                                    <asp:DropDownList ID="ddlEmpresaVenda" runat="server" Font-Names="Arial Narrow" 
                                                                        Width="547px">
                                                                    </asp:DropDownList>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td class="primario">
                                                                </td>
                                                                <td>
                                                                    <asp:CheckBox ID="chkConsolidarEmpresaVenda" runat="server" Text="Consolidar Empresa Venda" />
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td class="primario">
                                                                    Ano&#160;
                                                                </td>
                                                                <td>
                                                                    <asp:DropDownList ID="ddlAno" runat="server" Font-Names="Arial Narrow" 
                                                                        Width="76px">
                                                                    </asp:DropDownList>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td class="rotulo">
                                                                    Mercado
                                                                </td>
                                                                <td>
                                                                    <asp:RadioButton ID="rdInterno" runat="server" GroupName="mercado" Text="Interno" />&#160;&#160;<asp:RadioButton
                                                                        ID="rdExterno" runat="server" GroupName="mercado" Text="Externo" /><asp:RadioButton
                                                                            ID="rdTodosMercados" runat="server" Checked="True" GroupName="mercado" Text="Todos" />
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td class="primario">
                                                                </td>
                                                                <td>
                                                                    <asp:CheckBox ID="chkConsiderarCompras" runat="server" AutoPostBack="True" Checked="True"
                                                                        OnCheckedChanged="chkConsiderarCompras_CheckedChanged" Text="Considerar Compras" />
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td class="primario" colspan="2">
                                                                    <asp:Panel ID="pnlEmpresaCompra" runat="server">
                                                                        <table style="width: 100%;">
                                                                            <tr>
                                                                                <td>
                                                                                    Empresa Compra
                                                                                </td>
                                                                                <td>
                                                                                    <asp:DropDownList ID="ddlEmpresaCompra" runat="server" Font-Names="Arial Narrow"
                                                                                         Width="547px">
                                                                                    </asp:DropDownList>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </asp:Panel>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </div>
                                                </td>
                                                <td align="right" style="width: 5%" valign="middle">
                                                    <asp:Button ID="btnPremissas" runat="server" CssClass="botao" Text="Premissas" OnClick="btnPremissas_Click" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <uc:SelecaoProduto ID="ucSelecaoProdutoVendas" runat="server" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2">
                                                    <uc:SelecaoProduto ID="ucSelecaoProdutoCompras" runat="server" />
                                                </td>
                                            </tr>
                                        </table>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
                                    <HeaderTemplate>
                                        Premissas
                                    </HeaderTemplate>
                                    <ContentTemplate>
                                        <asp:Panel ID="Panel2" runat="server" Height="600px" Width="800px">
                                            <table style="width: 100%; border: 0px none;">
                                                <tr>
                                                    <td colspan="5" style="font-weight: bold; font-size: 22px; color: black; font-family: Arial;
                                                        background-color: darkgray; text-align: center">
                                                        Premissas do CRM
                                                    </td>
                                                    <td colspan="1" rowspan="18" valign="top">
                                                        <table>
                                                            <tr>
                                                                <td>
                                                                    <asp:Button ID="btnSelecaoClientes" runat="server" CssClass="botao" Text="Selecao Clientes"
                                                                        OnClick="btnSelecaoClientes_Click" />
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td>
                                                                    <asp:Button ID="btnLimparPremissas" runat="server" CssClass="botao" Text="Aplicar Padrão"
                                                                        OnClick="btnLimparPremissas_Click" />
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td>
                                                                    &#160;&#160;
                                                                </td>
                                                            </tr>
                                                        </table>
                                                        &#160; &#160; &#160; &#160; &#160; &#160;
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="rotulo" rowspan="1">
                                                        &#160;&#160;
                                                    </td>
                                                    <td rowspan="1" class="rotulo" align="center">
                                                        Faturamento Venda<br />
                                                        <br />
                                                        (R$/Ano)<br />
                                                        <br />
                                                        % Corte CRM
                                                        <asp:TextBox ID="txtPercentualCorte" runat="server" Width="60px" AutoPostBack="True"
                                                            onkeyup="formatavalor(this, 3, 0);" OnTextChanged="txtPercentualCorte_TextChanged">80</asp:TextBox>&nbsp;
                                                    </td>
                                                    <td align="right" class="rotulo" style="font-weight: bold; background-color: #d46b13;
                                                        text-align: center">
                                                        &#160;Cliente de Massa
                                                    </td>
                                                    <td class="rotulo" align="center">
                                                        &lt; &#160;<asp:Label ID="lblCorte" runat="server" Text="0,00"/>=&gt;
                                                    </td>
                                                    <td style="background-color: #ccccff; text-align: center" class="rotulo">
                                                        Cliente CRM
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                    </td>
                                                    <td>
                                                    </td>
                                                    <td>
                                                    </td>
                                                    <td>
                                                    </td>
                                                    <td>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="rotulo" style="height: 14px">
                                                    </td>
                                                    <td style="height: 14px" class="rotulo">
                                                    </td>
                                                    <td style="height: 14px" class="rotulo">
                                                    </td>
                                                    <td style="height: 14px; background-color: lightgrey" align="center" class="rotulo">
                                                        Escala
                                                    </td>
                                                    <td style="height: 14px; background-color: lightgrey" align="center" class="rotulo">
                                                        Pontos
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="center" class="rotulo" rowspan="3" style="background-color: lightgrey">
                                                        Peso
                                                        <asp:TextBox ID="txtPesoMargemBrutaVenda" runat="server" BackColor="#80FF80" Width="30px"
                                                            onkeyup="formatavalor(this, 3, 0);" AutoPostBack="True" OnTextChanged="txtPesoMargemBrutaVenda_TextChanged">34</asp:TextBox>%
                                                    </td>
                                                    <td align="center" class="primario" style="background-color: lightgrey">
                                                        &#160; Margem Bruta Venda (%/Ano)
                                                    </td>
                                                    <td align="right" class="rotulo" style="background-color: lightgrey">
                                                        &lt;
                                                    </td>
                                                    <td class="rotulo" style="background-color: lightgrey">
                                                        <asp:TextBox ID="txtMBVPercMenor" runat="server" Width="60px" AutoPostBack="True"
                                                            onkeyup="formatavalor(this, 2, 0);" OnTextChanged="txtMBVPercMenor_TextChanged">10</asp:TextBox>%
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: lightgrey">
                                                        <asp:Label ID="lblMBVMenor" runat="server" Text="3"/>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="center" class="rotulo" style="background-color: lightgrey">
                                                    </td>
                                                    <td align="right" class="rotulo" style="background-color: lightgrey">
                                                        Entre
                                                    </td>
                                                    <td class="rotulo" style="background-color: lightgrey">
                                                        <asp:TextBox ID="txtMBVPercDe" runat="server" Width="60px" onkeyup="formatavalor(this, 2, 0);"
                                                            Enabled="False">10</asp:TextBox>%&#160; A&#160;
                                                        <asp:TextBox ID="txtMBVPercAte" runat="server" Width="60px" onkeyup="formatavalor(this, 2, 0);"
                                                            Enabled="False">15</asp:TextBox>%
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: lightgrey">
                                                        <asp:Label ID="lblMBVEntre" runat="server" Text="6"/>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="left" class="rotulo" style="background-color: lightgrey">
                                                    </td>
                                                    <td align="right" class="rotulo" style="background-color: lightgrey">
                                                        &gt;
                                                    </td>
                                                    <td class="rotulo" style="background-color: lightgrey">
                                                        <asp:TextBox ID="txtMBVPercMaior" runat="server" Width="60px" AutoPostBack="True"
                                                            onkeyup="formatavalor(this, 2, 0);" OnTextChanged="txtMBVPercMaior_TextChanged">15</asp:TextBox>%
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: lightgrey">
                                                        <asp:Label ID="lblMBVMaior" runat="server" Text="9"/>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="left" class="rotulo" style="background-color: lightgrey">
                                                    </td>
                                                    <td align="left" class="rotulo" style="background-color: lightgrey">
                                                    </td>
                                                    <td align="right" class="rotulo" style="background-color: lightgrey">
                                                    </td>
                                                    <td class="rotulo" style="background-color: lightgrey">
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: lightgrey">
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="center" class="rotulo" rowspan="3">
                                                        Peso
                                                        <asp:TextBox ID="txtPesoParticipacaoSOC" runat="server" BackColor="#80FF80" Width="30px"
                                                            onkeyup="formatavalor(this, 5, 0);" AutoPostBack="True" OnTextChanged="txtPesoParticipacaoSOC_TextChanged">33</asp:TextBox>%
                                                    </td>
                                                    <td align="center" class="primario">
                                                        Participação/SOC %
                                                    </td>
                                                    <td align="right" class="rotulo">
                                                        &lt;
                                                    </td>
                                                    <td class="rotulo">
                                                        <asp:TextBox ID="txtSOCPercMenor" runat="server" Width="60px" AutoPostBack="True"
                                                            onkeyup="formatavalor(this, 2, 0);" OnTextChanged="txtSOCPercMenor_TextChanged">50</asp:TextBox>%
                                                    </td>
                                                    <td align="center" class="rotulo">
                                                        <asp:Label ID="lblSocMenor" runat="server" Text="3"/>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="center" class="rotulo">
                                                        Peso
                                                        <asp:TextBox ID="txtPesoVenda" runat="server" onkeyup="formatavalor(this, 3, 0);"
                                                            AutoPostBack="True" BackColor="#FFFF80" Width="30px" OnTextChanged="txtPesoVenda_TextChanged">70</asp:TextBox>&#160;%&#160;
                                                        Venda
                                                    </td>
                                                    <td align="right" class="rotulo">
                                                        Entre
                                                    </td>
                                                    <td class="rotulo">
                                                        <asp:TextBox ID="txtSOCPercDe" runat="server" Width="60px" onkeyup="formatavalor(this, 2, 0);"
                                                            Enabled="False">50</asp:TextBox>%&#160; A&#160;
                                                        <asp:TextBox ID="txtSOCPerAte" runat="server" Width="60px" onkeyup="formatavalor(this, 2, 0);"
                                                            Enabled="False">70</asp:TextBox>%
                                                    </td>
                                                    <td align="center" class="rotulo">
                                                        <asp:Label ID="lblSocEntre" runat="server" Text="6"/>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="center" class="rotulo">
                                                        &#160; Peso
                                                        <asp:TextBox ID="txtPesoCompra" runat="server" onkeyup="formatavalor(this, 3, 0);"
                                                            BackColor="#FFFF80" Width="30px" OnTextChanged="txtPesoCompra_TextChanged">30</asp:TextBox>&#160;%
                                                        Compra
                                                    </td>
                                                    <td align="right" class="rotulo">
                                                        &gt;
                                                    </td>
                                                    <td class="rotulo">
                                                        <asp:TextBox ID="txtSOCPercMaior" runat="server" Width="60px" AutoPostBack="True"
                                                            onkeyup="formatavalor(this, 2, 0);" OnTextChanged="txtSOCPercMaior_TextChanged">70</asp:TextBox>%
                                                    </td>
                                                    <td align="center" class="rotulo">
                                                        <asp:Label ID="lblSocMaior" runat="server" Text="9"/>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="left" class="rotulo">
                                                    </td>
                                                    <td align="left" class="rotulo">
                                                    </td>
                                                    <td align="right" class="rotulo">
                                                    </td>
                                                    <td class="rotulo">
                                                    </td>
                                                    <td align="center" class="rotulo">
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="center" class="rotulo" rowspan="3" style="background-color: lightgrey">
                                                        Peso
                                                        <asp:TextBox ID="txtPesoRatingDeCredito" runat="server" BackColor="#80FF80" Width="30px"
                                                            onkeyup="formatavalor(this, 3, 0);" AutoPostBack="True" OnTextChanged="txtPesoRatingDeCredito_TextChanged">33</asp:TextBox>%
                                                    </td>
                                                    <td align="center" class="primario" style="background-color: lightgrey">
                                                        Rating De Credito
                                                    </td>
                                                    <td align="right" class="rotulo" style="background-color: lightgrey">
                                                        C
                                                    </td>
                                                    <td class="rotulo" style="background-color: lightgrey">
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: lightgrey">
                                                        <asp:Label ID="lblRatingC" runat="server" Text="3"/>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style="background-color: lightgrey" align="center" class="rotulo">
                                                    </td>
                                                    <td style="background-color: lightgrey" align="right" class="rotulo">
                                                        B
                                                    </td>
                                                    <td style="background-color: lightgrey" class="rotulo">
                                                    </td>
                                                    <td style="background-color: lightgrey" align="center" class="rotulo">
                                                        <asp:Label ID="lblRatingB" runat="server" Text="6"/>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="left" class="rotulo" style="background-color: lightgrey">
                                                    </td>
                                                    <td align="right" class="rotulo" style="background-color: lightgrey">
                                                        A
                                                    </td>
                                                    <td class="rotulo" style="background-color: lightgrey">
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: lightgrey">
                                                        <asp:Label ID="lblRatingA" runat="server" Text="9"/>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                    </td>
                                                    <td>
                                                    </td>
                                                    <td>
                                                    </td>
                                                    <td>
                                                    </td>
                                                    <td>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="center" class="rotulo">
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: #003366; color: white;">
                                                        CLASSIFICACAO
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: #003366; color: white;">
                                                        BRONZE
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: #003366; color: white;">
                                                        PRATA
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: #003366; color: white;">
                                                        OURO
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="center" class="rotulo">
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: #003366; color: white;">
                                                        PONTOS
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: lightgrey">
                                                        &lt; OU = 18
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: lightgrey">
                                                        18,01 - 23,99
                                                    </td>
                                                    <td align="center" class="rotulo" style="background-color: lightgrey">
                                                        &gt; OU = 24
                                                    </td>
                                                </tr>
                                            </table>
                                        </asp:Panel>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                                <ajaxToolkit:TabPanel ID="TabPanel3" runat="server" HeaderText="TabPanel3">
                                    <HeaderTemplate>
                                        Seleção dos Clientes
                                    </HeaderTemplate>
                                    <ContentTemplate>
                                        <asp:Panel ID="Panel3" runat="server" Height="600px" Width="100%" ScrollBars="Vertical">
                                            <table style="width: 100%; border: 0px none">
                                                <tr>
                                                    <td>
                                                        <asp:Button ID="btnSalvar" runat="server" CssClass="botao" Text="Salvar" />
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        <asp:GridView ID="gridSelecaoClientes" runat="server" AutoGenerateColumns="False"
                                                            CellPadding="4" Font-Names="Arial Narrow" Font-Size="7pt" ForeColor="#333333"
                                                            GridLines="None" Width="100%">
                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                            <Columns>
                                                                <asp:BoundField DataField="PercentualFaturamentoAcumulado" DataFormatString="{0:N2}"
                                                                    HeaderText="Fat. Acum." />
                                                                <asp:TemplateField HeaderText="Tipo">
                                                                    <ItemTemplate>
                                                                        <asp:DropDownList ID="ddlTipo" runat="server" Font-Names="Arial Narrow" 
                                                                            Height="16px" SelectedValue='<%# eval("TipoClienteCRM") %>' Width="85px" AutoPostBack="True"
                                                                            OnSelectedIndexChanged="ddlTipo_SelectedIndexChanged">
                                                                            <asp:ListItem Value="V">-</asp:ListItem>
                                                                            <asp:ListItem Value="C">CRM</asp:ListItem>
                                                                            <asp:ListItem Value="P">Prospects</asp:ListItem>
                                                                            <asp:ListItem Value="E">Estrategico</asp:ListItem>
                                                                            <asp:ListItem Value="M">Massa</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </ItemTemplate>
                                                                </asp:TemplateField>
                                                                <asp:TemplateField HeaderText="Qualitativo">
                                                                    <ItemTemplate>
                                                                        <asp:DropDownList ID="ddlQualitativo" runat="server" Font-Names="Arial Narrow" 
                                                                            Height="16px" Width="101px" AutoPostBack="True" OnSelectedIndexChanged="ddlQualitativo_SelectedIndexChanged"
                                                                            SelectedValue='<%# eval("TipoClienteQualitativo") %>'>
                                                                            <asp:ListItem Value="V">-</asp:ListItem>
                                                                            <asp:ListItem Value="N">Negocio</asp:ListItem>
                                                                            <asp:ListItem Value="R">Relacionamento</asp:ListItem>
                                                                            <asp:ListItem Value="P">Preco</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </ItemTemplate>
                                                                </asp:TemplateField>
                                                                <asp:BoundField DataField="CodigoCliente" HeaderText="Cliente" />
                                                                <asp:BoundField DataField="NomeCliente" HeaderText="NomeCliente" />
                                                                <asp:BoundField DataField="PercentualMargemBruta" DataFormatString="{0:N2}" HeaderText="% Margem Bruta Venda" />
                                                                <asp:BoundField DataField="PontuacaoMBV" HeaderText="Pontuacao MBV" />
                                                                <asp:BoundField DataField="PotencialVenda" DataFormatString="{0:N2}" HeaderText="Potencial Venda" />
                                                                <asp:BoundField DataField="TotalVendas" DataFormatString="{0:N2}" HeaderText="Total Vendas" />
                                                                <asp:BoundField DataField="PercentualVendas" DataFormatString="{0:N2}" HeaderText="% Vendas" />
                                                                <asp:BoundField DataField="PotencialCompra" DataFormatString="{0:N2}" HeaderText="Potencial Compra" />
                                                                <asp:BoundField DataField="TotalCompras" DataFormatString="{0:N2}" HeaderText="Total Compras" />
                                                                <asp:BoundField DataField="PercentualCompras" DataFormatString="{0:N2}" HeaderText="% Compras" />
                                                                <asp:BoundField DataField="PontuacaoSOC" HeaderText="PontuacaoSOC" />
                                                                <asp:BoundField DataField="Agrupamento" HeaderText="Agrup. Cred." />
                                                                <asp:BoundField DataField="LimiteCredito" HeaderText="Rating" />
                                                                <asp:BoundField DataField="PontuacaoRating" HeaderText="PontuacaoRating" />
                                                                <asp:BoundField DataField="Classificacao" HeaderText="Classificacao" />
                                                            </Columns>
                                                            <EditRowStyle BackColor="#999999" />
                                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                        </asp:GridView>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        &#160;&#160;
                                                    </td>
                                                </tr>
                                            </table>
                                        </asp:Panel>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                            </ajaxToolkit:TabContainer>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
