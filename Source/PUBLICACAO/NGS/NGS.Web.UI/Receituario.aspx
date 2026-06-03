<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Receituario.aspx.vb" Inherits="NGS.Web.UI.Receituario" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngReceituario" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlReceituario" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Receituário Agronômico
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
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
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                    <asp:TextBox ID="txtEmpresa" runat="server" Width="550px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnEmpresa" OnClick="btnEmpresa_Click" runat="server" Enabled="False"
                        CssClass="btn" UseSubmitBehavior="False" Text=">" data-ToolTip="default"
                        ToolTip="Empresa Para negociação/consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" Width="90px" CssClass="calendario" runat="server"
                        data-ToolTip="default" ToolTip="Período do receituário." />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" Width="90px" CssClass="calendario" runat="server"
                        data-ToolTip="default" ToolTip="Período do receituário." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="ImgConsultar" OnClick="ImgConsultar_Click" runat="server" Width="22px"
                        Height="22px" ImageUrl="~/images/marcar.ico" Enabled="False" ImageAlign="AbsMiddle"
                        data-ToolTip="default" ToolTip="Consultar" Visible="false" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="ImgLimpar" OnClick="ImgLimpar_Click" runat="server" Width="22px"
                        Height="22px" ImageUrl="~/images/desmarcar.ico" Enabled="False" ImageAlign="AbsMiddle"
                        data-ToolTip="default" ToolTip="Limpar" Visible="false" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="ImgSair" runat="server" Width="22px" Height="22px" ImageUrl="~/Images/Sair.jpg"
                        ImageAlign="AbsMiddle" data-ToolTip="default" ToolTip="Sair" OnClientClick="MenuDeAcesso('Revenda');"
                        Visible="false" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Reimpressão:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="radSim" runat="server" GroupName="Reimpressao" Text="Sim "
                        data-ToolTip="default" ToolTip="Selecionar a opção desejada." />
                    <asp:RadioButton ID="radNao" runat="server" GroupName="Reimpressao" Text="Não "
                        data-ToolTip="default" ToolTip="Selecionar a opção desejada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nota Fiscal:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNota" runat="server" data-ToolTip="default" ToolTip="Número da nota fiscal." />
                </div>
            </div>
            <ajaxToolkit:TabContainer ID="TabReceituario" runat="server" Width="100%" ActiveTabIndex="1"
                Style="margin-top: 4px;">
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel1" ID="TabPanel1">
                    <HeaderTemplate>
                        Notas Fiscais
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid" style="height: 300px;">
                            <asp:GridView ID="gridNotas" runat="server" Width="100%" OnSelectedIndexChanged="gridNotas_SelectedIndexChanged"
                                ForeColor="#333333" GridLines="None" CellPadding="4" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                                    <asp:BoundField DataField="Nota" HeaderText="Nota" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoCliente" HeaderText="Cliente" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EndCliente" HeaderText="End" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nome" HeaderText="Nome" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Complemento" HeaderText="Complemento" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
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
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel2" ID="TabPanel2">
                    <HeaderTemplate>
                        Itens da Nota Fiscal
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid" style="height: 200px;">
                            <asp:GridView ID="gridItem" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridItem_SelectedIndexChanged"
                                OnRowDataBound="gridItem_RowDataBound">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>                                            
                                            <asp:LinkButton ID="lnkSelecionar" CssClass="lnk" 
                                                data-tooltip="default" ToolTip="Selecionar registro." runat="server" Text=" &gt; "
                                                CommandName="Select">
                                                <i class="fa fa-arrow-right seta"></i>
                                            </asp:LinkButton>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nome" HeaderText="Nome" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Unidade" HeaderText="Unidade" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoReceita" HeaderText="Receita" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:LinkButton Text="Imprimir" ID="lnkImprimir" runat="server" OnClick="lnkImprimir_Click"></asp:LinkButton>
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:TemplateField>
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Agrônomo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlAgronomos" runat="server" Width="550px" OnSelectedIndexChanged="ddlAgronomos_SelectedIndexChanged"
                                    AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                ART:
                            </div>
                            <div class="coltxt" style="width: 86.13%;">
                                <div class="bordagrid" style="height: 50px;">
                                    <asp:GridView ID="gridART" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                        ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridART_SelectedIndexChanged">
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                        <EditRowStyle BackColor="#999999" />
                                        <Columns>
                                            <asp:CommandField ButtonType="Button" SelectText=" &gt;" ShowSelectButton="True" />
                                            <asp:BoundField DataField="CodigoArt" HeaderText="ART" HtmlEncode="False">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ARTInicial" HeaderText="Inicial" HtmlEncode="False">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ARTFinal" HeaderText="Final" HtmlEncode="False">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ArtAtual" HeaderText="Emitidas" HtmlEncode="False">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Status" HeaderText="Status" HtmlEncode="False">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                        </Columns>
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    </asp:GridView>
                                </div>
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:ImageButton ID="imgEmitirReceituario" runat="server" ImageUrl="~/images/confirmar.gif"
                                    OnClick="imgEmitirReceituario_Click" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel3" ID="TabPanel3">
                    <HeaderTemplate>
                        Receita
                    </HeaderTemplate>
                    <ContentTemplate>
                        <table style="border-top-style: none; border-right-style: none; border-left-style: none; border-bottom-style: none; width: 100%; border: 0px none;">
                            <tr>
                                <td class="bordatotal" colspan="2">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">CULTURA</span></strong></label>&nbsp;<br />
                                    <asp:DropDownList ID="ddlCultura" runat="server" Width="500px" OnSelectedIndexChanged="ddlCultura_SelectedIndexChanged"
                                        AutoPostBack="True">
                                    </asp:DropDownList>
                                </td>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">FITOSSANITÁRIO</span></strong></label>&nbsp;<br />
                                    <asp:Label ID="lblFito" runat="server" Font-Bold="False" ForeColor="Blue" BorderStyle="None" />
                                </td>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">FORMULAÇÃO</span></strong></label>&nbsp;<br />
                                    <asp:Label ID="lblFormulacaoFito" runat="server" Font-Bold="False" ForeColor="Blue"
                                        BorderStyle="None" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">PRAGA</span></strong></label>&nbsp;<br />
                                    <asp:DropDownList ID="ddlPraga" runat="server" Width="230px" OnSelectedIndexChanged="ddlPraga_SelectedIndexChanged"
                                        AutoPostBack="True">
                                    </asp:DropDownList>
                                </td>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">TIPO DE APLICAÇÃO</span></strong></label>&nbsp;<br />
                                    <asp:DropDownList ID="ddlTipoDeAplicacao" runat="server" Width="230px"
                                        OnSelectedIndexChanged="ddlTipoDeAplicacao_SelectedIndexChanged"
                                        AutoPostBack="True">
                                    </asp:DropDownList>
                                </td>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">NOME COMERCIAL</span></strong></label>&nbsp;<br />
                                    <asp:Label ID="lblNomeComercial" runat="server" Font-Bold="False" ForeColor="Blue"
                                        BorderStyle="None" />
                                </td>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">NOME TÉCNICO</span></strong></label>&nbsp;<br />
                                    <asp:Label ID="lblNomeTecnico" runat="server" Font-Bold="False" ForeColor="Blue"
                                        BorderStyle="None" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bordatotal" colspan="2" rowspan="4">
                                    <asp:Panel ID="pnlDosagem" runat="server" Width="500px" Height="100px" BorderStyle="Double"
                                        BorderColor="#336699" ScrollBars="Vertical">
                                        <asp:GridView ID="gridDosagem" runat="server" Width="98%" OnSelectedIndexChanged="gridDosagem_SelectedIndexChanged"
                                            ForeColor="#333333" GridLines="None" CellPadding="4" AutoGenerateColumns="False">
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                                            <Columns>
                                                <asp:CommandField SelectText=" &gt;" ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                                                <asp:BoundField DataField="CodigoDosagem" HeaderText="Dosagem">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="CodigoSolo" HeaderText="Solo">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="UnidadeDeMedida" HeaderText="UND. de Medida">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="DosagemMinima" HeaderText="M&#237;nima">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="DosagemMaxima" HeaderText="M&#225;xima">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="DosagemRecomendada" HeaderText="Recomendado">
                                                    <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundField>
                                            </Columns>
                                            <EditRowStyle BackColor="#999999"></EditRowStyle>
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></FooterStyle>
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></HeaderStyle>
                                            <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White"></PagerStyle>
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333"></RowStyle>
                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>
                                        </asp:GridView>
                                    </asp:Panel>
                                    <table style="width: 505px; border: 0px none;">
                                        <tr>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">QUANTIDADE DO PRODUTO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="lblQuantidadeDoProduto" runat="server" Font-Bold="False" ForeColor="Blue"
                                                    BorderStyle="None" />
                                            </td>
                                            <td class="primario">
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">NR.APLICAÇÕES</span></strong></label>&nbsp;<br />
                                                <asp:TextBox ID="txtNumeroAplicacao" runat="server" Width="30px" AutoPostBack="True"
                                                    ForeColor="Red" BorderStyle="None" OnTextChanged="txtNumeroAplicacao_TextChanged" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">DOSAGEM MÍNIMA</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="lblDosagemMinima" runat="server" Font-Bold="False" ForeColor="Blue"
                                                    BorderStyle="None" />
                                            </td>
                                            <td class="primario">
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">RECOMENDADO</span></strong></label>&nbsp;<br />
                                                <asp:TextBox ID="txtDosagemRecomendada" CssClass="txtDecimal4" runat="server" Width="110px"
                                                    AutoPostBack="True" ForeColor="Red" BorderStyle="None" OnTextChanged="txtDosagemRecomendada_TextChanged" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">DOSAGEM MÁXIMA</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="lblDosagemMaxima" runat="server" Font-Bold="False" ForeColor="Blue"
                                                    BorderStyle="None" />
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">CLASSE TOXICOLÓGICA</span></strong></label>&nbsp;<br />
                                    <asp:Label ID="lblClasseToxicologica" runat="server" Font-Bold="False" ForeColor="Blue"
                                        BorderStyle="None" />
                                </td>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">CLASSE AMBIENTAL</span></strong></label>&nbsp;<br />
                                    <asp:Label ID="lblClasseAmbiental" runat="server" Font-Bold="False" ForeColor="Blue"
                                        BorderStyle="None" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">VAZÃO</span></strong></label>&nbsp;<br />
                                    <asp:Label ID="lblVazao" runat="server" Font-Bold="False" ForeColor="Blue" BorderStyle="None" />
                                </td>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">CLASSE DE RISCO</span></strong></label>&nbsp;<br />
                                    <asp:Label ID="lblClasseDeRisco" runat="server" Font-Bold="False" ForeColor="Blue"
                                        BorderStyle="None" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">ÁREA TRATADA</span></strong></label>&nbsp;<br />
                                    <asp:Label ID="lblAreaTratada" runat="server" Font-Bold="False" ForeColor="Blue"
                                        BorderStyle="None" />
                                </td>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt; color: red">ÁREA TOTAL</span></strong></label>&nbsp;<br />
                                    <asp:TextBox ID="txtAreaPlantada" CssClass="txtDecimal" runat="server" Width="100px"
                                        ForeColor="Red" BorderStyle="None" BackColor="White" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">INTERVALO DE SEGURANÇA</span></strong></label>&nbsp;<br />
                                    <asp:Label ID="lblIntervaloDeSeguranca" runat="server" Font-Bold="False" ForeColor="Blue"
                                        BorderStyle="None" />
                                </td>
                                <td class="bordatotal">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">LOCAL DE APLICAÇÃO</span></strong></label>&nbsp;<br />
                                    <asp:Label ID="lblLocalDeAplicacao" runat="server" Font-Bold="False" ForeColor="Blue"
                                        BorderStyle="None" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bordatotal" colspan="2">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">INSTRUÇÕES DE MANEJO</span></strong></label>&nbsp;<br />
                                    <asp:TextBox ID="txtInstrucaoDeManejo" runat="server" Width="500px" Height="80px"
                                        ReadOnly="True" TextMode="MultiLine" />
                                </td>
                                <td class="bordatotal" colspan="2">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">ÉPOCA DE APLICAÇÃO</span></strong></label>&nbsp;<br />
                                    <asp:TextBox ID="txtEpocaDeAplicacao" runat="server" Width="500px" Height="80px"
                                        ReadOnly="True" TextMode="MultiLine" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bordatotal" colspan="2">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">MODALIDADE DE APLICAÇÃO</span></strong></label>&nbsp;<br />
                                    <asp:TextBox ID="txtModalidadeDeAplicacao" runat="server" Width="500px" Height="80px"
                                        ReadOnly="True" TextMode="MultiLine" />
                                </td>
                                <td class="bordatotal" colspan="2">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">PRIMEIRO SOCORROS</span></strong></label>&nbsp;<br />
                                    <asp:TextBox ID="txtPrimeiroSocorros" runat="server" Width="500px" Height="80px"
                                        ReadOnly="True" TextMode="MultiLine" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bordatotal" colspan="2">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">ADVERTÊNCIAS</span></strong></label>&nbsp;<br />
                                    <asp:TextBox ID="txtAdvertenciaMeioAmbiente" runat="server" Width="500px" Height="80px"
                                        ReadOnly="True" TextMode="MultiLine" />
                                </td>
                                <td class="bordatotal" colspan="2">
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">INSTRUÇÕES EMBALAGEM</span></strong></label>&nbsp;<br />
                                    <asp:TextBox ID="txtInstrucaoEmbalagem" runat="server" Width="500px" Height="80px"
                                        ReadOnly="True" TextMode="MultiLine" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bordatotal" colspan="4">
                                    <asp:ImageButton ID="imgConfirmaReceita" OnClick="imgConfirmaReceita_Click" runat="server"
                                        ImageUrl="~/images/confirmar.gif"></asp:ImageButton>
                                </td>
                            </tr>
                        </table>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
</asp:Content>
