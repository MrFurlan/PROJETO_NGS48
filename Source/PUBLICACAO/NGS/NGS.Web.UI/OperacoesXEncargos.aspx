<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="OperacoesXEncargos.aspx.vb" Inherits="NGS.Web.UI.OperacoesXEncargos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1200px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmOperacoesXEncargos" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlOperacoesXEncargos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Operações X Encargos
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Registro
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton Text="Limpar" ID="lnkLimpar" runat="server" />
                                    </li>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkPlanoDeContas" Text="Plano de Contas" runat="server" />
                                    </li>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div>
                            <div class="painelleft" style="width: 66%;">
                                <div class="row">
                                    <div class="collbl">
                                        ID Config.:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtCodigoConfig" runat="server" Width="50px" CssClass="txtNumerico9" AutoPostBack="True"
                                            Style="text-align: right; background-color: greenyellow;" />
                                    </div>
                                    <div class="collbl">
                                        Inicio Vigencia:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtVigencia" runat="server" class="calendario" />
                                    </div>
                                    <div class="coltxt" style="margin-left: 40px">
                                        <asp:CheckBox ID="chkAtivo" runat="server" AutoPostBack="True" Text="Ativo"
                                            OnCheckedChanged="Ativo_CheckedChanged" Enabled="false" data-ToolTip="default"
                                            ToolTip="Desativar ou ativar ID." />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Empresa:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlEmpresa" runat="server" Width="500px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Grupo:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlGruposDeEstoques" runat="server" Width="500px" AutoPostBack="True"
                                            OnSelectedIndexChanged="DdlGruposDeEstoques_SelectedIndexChanged" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Produto:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlProdutos" runat="server" Width="500px" AutoPostBack="True"
                                            OnSelectedIndexChanged="DdlProdutos_SelectedIndexChanged" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Operação:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlOperacoes" runat="server" AutoPostBack="True" Width="251px"
                                            OnSelectedIndexChanged="DdlOperacoes_SelectedIndexChanged" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlSubOperacoes" runat="server" Width="245px" AutoPostBack="True"
                                            OnSelectedIndexChanged="DdlSubOperacoes_SelectedIndexChanged" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        UF Orig/Dest:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlEstadoOrigem" runat="server" Width="251px" AutoPostBack="True"
                                            OnSelectedIndexChanged="DdlEstadoOrigem_SelectedIndexChanged" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlEstadoDestino" runat="server" Width="245px" AutoPostBack="True"
                                            OnSelectedIndexChanged="DdlEstadoDestino_SelectedIndexChanged" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Grupo CFOP:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlCFOPTitulos" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DdlCFOPTitulos_SelectedIndexChanged"
                                            Width="500px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        CFOP:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlCFOP" runat="server" Width="500px" AutoPostBack="True" OnSelectedIndexChanged="DdlCFOP_SelectedIndexChanged" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Sit. ICMS:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlSituacaoTributariaICMS" runat="server" Width="500px" AutoPostBack="True"
                                            OnSelectedIndexChanged="ddlSituacaoTributariaICMS_SelectedIndexChanged" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        OBS. ICMS:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlObsICMS" runat="server" Width="500px" AutoPostBack="True" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        BENEFÍCIO ICMS:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlBeneficioICMS" runat="server" Width="500px" OnSelectedIndexChanged="ddlBeneficioICMS_SelectedIndexChanged" AutoPostBack="True" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Sit. IPI:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlSituacaoTributariaIPI" runat="server" Width="500px" AutoPostBack="True"
                                            OnSelectedIndexChanged="ddlSituacaoTributariaIPI_SelectedIndexChanged" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Sit. PIS/COFINS:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlSituacaoTributariaPISCOFINS" runat="server" Width="500px"
                                            AutoPostBack="True" OnSelectedIndexChanged="ddlSituacaoTributariaPISCOFINS_SelectedIndexChanged" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Obs. PIS/COFINS:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlObsPISCOFINS" runat="server" Width="500px" AutoPostBack="True"
                                            OnSelectedIndexChanged="ddlObsPISCOFINS_SelectedIndexChanged" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Sit. IBS/CBS:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlSituacaoTributariaIBSCBS" runat="server" Width="500px" AutoPostBack="True"
                                            OnSelectedIndexChanged="ddlSituacaoTributariaIBSCBS_SelectedIndexChanged" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Clas. Trib. IBS\CBS:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlClassificacaoIBSCBS" runat="server" Width="500px" AutoPostBack="True" OnSelectedIndexChanged="ddlClassificacaoIBSCBS_SelectedIndexChanged" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="coltxt">
                                        <asp:CheckBox ID="chkNaturezaDeRendimento" ClientIDMode="Static" runat="server" AutoPostBack="True" OnCheckedChanged="chkNaturezaDeRendimento_CheckedChanged" Text="Informar Natureza De Rendimento"
                                            Font-Bold="True" data-ToolTip="default" ToolTip="Selecionar Natureza De Rendimento para informação no SPED REINF." />
                                    </div>
                                    <div class="row" id="divNaturezaDeRendimento" runat="server" visible="false">
                                        <div class="collbl colw">
                                            N.Rendimento:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlNaturezaDeRendimento" runat="server" Width="500px" AutoPostBack="True"
                                                OnSelectedIndexChanged="ddlNaturezaDeRendimento_SelectedIndexChanged" />
                                        </div>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="coltxt">
                                        <asp:CheckBox ID="chkCalculadora" ClientIDMode="Static" runat="server" AutoPostBack="True" OnCheckedChanged="chkCalculadora_CheckedChanged" Text="Usar a calculadora de impostos Sefaz"
                                            Font-Bold="True" data-ToolTip="default" ToolTip="Usar a calculadora de impostos Sefaz" />
                                    </div>
                                </div>
                            </div>
                            <div class="painelright" style="width: 33%;">
                                <div class="subtitulodiv">
                                    Inicio Vigencia / Versão / Usuario / Data Inclusão
                                </div>
                                <div class="row">
                                    <asp:DropDownList ID="ddlVersao" runat="server" Width="100%" AutoPostBack="True" />
                                </div>
                                <div class="subtitulodiv">
                                    Encargos
                                </div>
                                <div class="bordagrid" style="height: 282px;">
                                    <asp:GridView ID="gridEncargosAdd" runat="server" CellPadding="4" ForeColor="#333333"
                                        GridLines="None" AutoGenerateColumns="False" Width="100%" ShowHeader="False">
                                        <EditRowStyle BackColor="#999999" />
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:TemplateField>
                                                <ItemTemplate>
                                                    <asp:Button ID="BtnAdd" runat="server" Text=" + " OnClick="BtnAdd_Click" />
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="Encargo" HeaderText="Encargo" />
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </div>
                        </div>
                        <div class="bordagrid" style="height: auto;">
                            <asp:GridView ID="GridEncargos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnRowDataBound="GridEncargos_RowDataBound">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField HeaderText="Encargo">
                                        <ItemTemplate>
                                            <asp:Label ID="lblEncargo" runat="server" Text='<%# Bind("CodigoEncargo") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Débito">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtDebito" runat="server" MaxLength="9" Width="58px"
                                                data-tooltip="default" OnTextChanged="txtDebito_TextChanged" AutoPostBack="true" />
                                            <asp:ImageButton ID="imbContaDebito" runat="server" Style="cursor: pointer; border: 0 none;"
                                                data-ToolTip="default" ToolTip="Consulta de Plano de Contas" ImageAlign="AbsMiddle"
                                                ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imbContaDebito_click" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Crédito">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtCredito" runat="server" MaxLength="9" Width="58px" data-tooltip="default" OnTextChanged="txtCredito_TextChanged" AutoPostBack="true" />
                                            <asp:ImageButton ID="imbContaCredito" runat="server" Style="cursor: pointer; border: 0 none;"
                                                data-ToolTip="default" ToolTip="Consulta de Plano de Contas" ImageAlign="AbsMiddle"
                                                ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imbContaCredito_click" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Sinal">
                                        <ItemTemplate>
                                            <asp:DropDownList ID="DdlSinal" runat="server" Width="82px" OnSelectedIndexChanged="DdlSinal_SelectedIndexChanged" AutoPostBack="true">
                                                <asp:ListItem Value=""></asp:ListItem>
                                                <asp:ListItem Value="=">= Igual</asp:ListItem>
                                                <asp:ListItem Value="+">+ Mais</asp:ListItem>
                                                <asp:ListItem Value="-">- Menos</asp:ListItem>
                                            </asp:DropDownList>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Sobre">
                                        <ItemTemplate>
                                            <asp:Label ID="lblValorOuPeso" runat="server" Text='<%# Bind("Encargo.ValorOuPeso") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="% Base">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtBase" runat="server" MaxLength="9" Width="80px" CssClass="txtDecimal9" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Aliquota">
                                        <ItemTemplate>
                                            <asp:TextBox ID="TxtAliquota" runat="server" MaxLength="9" Width="80px" CssClass="txtDecimal9" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Aliq.Exib.">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtAliqExib" runat="server" MaxLength="3" Width="80px" CssClass="txtDecimal" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Limite">
                                        <ItemTemplate>
                                            <asp:TextBox ID="TxtLimite" runat="server" MaxLength="3" Width="45px" CssClass="txtDecimal" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Observações">
                                        <ItemTemplate>
                                            <asp:DropDownList ID="ddlObservacao" runat="server" Width="200px" OnSelectedIndexChanged="ddlObservacao_SelectedIndexChanged" AutoPostBack="true" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluir" runat="server" data-ToolTip="default" ImageUrl="~/Images/deletar.gif"
                                                OnClick="imgExcluir_Click" OnClientClick="return confirm('Deseja realmente excluir o Encargo?');"
                                                Style="border: 0;" ToolTip="Excluir" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="GridConsulta" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridConsulta_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:BoundField DataField="Codigo_Id" HeaderText="ID" />
                                    <asp:BoundField DataField="InicioVigencia" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Inicio Vigencia" />
                                    <asp:BoundField DataField="Grupo" HeaderText="Grupo" HtmlEncode="False" />
                                    <asp:BoundField DataField="NomeGrupo" HeaderText="Descri&#231;&#227;o" HtmlEncode="False" />
                                    <asp:BoundField DataField="Produto" HeaderText="Produto" HtmlEncode="False" />
                                    <asp:BoundField DataField="NomeProduto" HeaderText="Descri&#231;&#227;o" HtmlEncode="False" />
                                    <asp:BoundField DataField="Operacao" HeaderText="Op" HtmlEncode="False" />
                                    <asp:BoundField DataField="SubOperacao" HeaderText="So" HtmlEncode="False" />
                                    <asp:BoundField DataField="NomeOperacao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False" />
                                    <asp:BoundField DataField="Origem" HeaderText="UF" HtmlEncode="False" />
                                    <asp:BoundField DataField="Destino" HeaderText="UF" HtmlEncode="False" />
                                    <asp:BoundField DataField="GrupoFiscal" HeaderText="Grp Fiscal" />
                                    <asp:BoundField DataField="CFOP" HeaderText="Cfop" HtmlEncode="False" />
                                    <asp:BoundField DataField="Ativo" HeaderText="Ativo">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel3" runat="server">
                    <HeaderTemplate>
                        Recontabilização
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultarNotas" Text="Consultar" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton Text="Limpar" ID="lnkLimparRec" runat="server" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRecontabilizar" Text="Recontabilizar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Id Versão:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtIdVersao" runat="server" Style="text-align: right;" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Periodo Inclusao:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPeriodoInicial" runat="server" CssClass="calendario" />
                                a
                                <asp:TextBox ID="txtPeriodoFinal" runat="server" CssClass="calendario" />
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            Versões Para Recontabilizar.
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Vesões:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlVersaoRec" runat="server" Width="500px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="bordagrid">
                                <asp:GridView ID="gridNotas" runat="server" CellPadding="4" ForeColor="#333333" GridLines="None"
                                    AutoGenerateColumns="False">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:TemplateField>
                                            <HeaderTemplate>
                                                <asp:CheckBox ID="chkNotaAll" runat="server" AutoPostBack="True" OnCheckedChanged="chkNotaAll_CheckedChanged" />
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:CheckBox ID="chkNota" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="OperacaoXEstado" HeaderText="Versão" />
                                        <asp:BoundField DataField="Empresa_id" HeaderText="Empresa" />
                                        <asp:BoundField DataField="Cliente_Id" HeaderText="Cliente" />
                                        <asp:BoundField DataField="EndCliente_id" HeaderText="End" />
                                        <asp:BoundField DataField="Nome" HeaderText="Nome" />
                                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento" />
                                        <asp:BoundField DataField="Nota_id" HeaderText="Nota" />
                                        <asp:BoundField DataField="Serie_id" HeaderText="Serie" />
                                        <asp:BoundField DataField="EntradaSaida_Id" HeaderText="E/S" />
                                        <asp:BoundField DataField="Produto_id" HeaderText="Código" />
                                        <asp:BoundField DataField="NomeProduto" HeaderText="Produto" />
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
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
    <uc:ConsultaObservacoes ID="ucConsultaObservacoes" runat="server" />
</asp:Content>
