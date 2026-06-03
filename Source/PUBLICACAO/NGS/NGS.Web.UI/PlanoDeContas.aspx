<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PlanoDeContas.aspx.vb" Inherits="NGS.Web.UI.PlanoDeContas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1180px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPlanoDeContas" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPlanoDeContas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Plano de Contas
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TaCadastro" runat="server" HeaderText="Cadastro">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                                        <ul>
                                            <li class="iconPdf">
                                                <asp:LinkButton ID="lnkPdf" runat="server" Text="Pdf" />
                                            </li>
                                            <li class="iconWord">
                                                <asp:LinkButton ID="lnkWord" runat="server" Text="Word" />
                                            </li>
                                            <li class="iconExcel">
                                                <asp:LinkButton ID="lnkExcel" runat="server" Text="Excel" />
                                            </li>
                                        </ul>
                                    </li>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkAdicionarContas" Text="Adicionar Contas" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="painelright" style="width: 35%;">
                            <asp:Panel ID="pnlEncargos" runat="server" Visible="False" Width="100%">
                                <div class="subtitulodiv">
                                    Contas Relacionadas:
                                </div>
                                <div class="bordagrid" style="height: 360px;">
                                    <asp:GridView ID="gridEncargos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                        ForeColor="#333333" GridLines="None" Width="100%">
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <EditRowStyle BackColor="#999999" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:BoundField DataField="CodigoContaEncargo" HeaderText="Conta">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TituloEncargo" HeaderText="Titulo">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:TemplateField>
                                                <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                                <ItemStyle HorizontalAlign="Left" Width="30px" />
                                                <ItemTemplate>
                                                    <asp:ImageButton ID="imgExcluirEncargo" runat="server" Height="18px" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_excluir.png"
                                                        OnClick="imgExcluirEncargo_Click" Style="cursor: pointer" data-ToolTip="default"
                                                        ToolTip="Excluir Encargo" Width="18px" />
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </asp:Panel>
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    Conta:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtConta" runat="server" MaxLength="9" CssClass="txtNumerico" Width="80px"
                                        data-ToolTip="default" ToolTip="Número da conta." />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtTitulo" runat="server" MaxLength="80" Width="288px"
                                        data-ToolTip="default" ToolTip="Nome da conta." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Tipo da Conta:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlTipoDaConta" runat="server" Width="392px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Tipo de Cliente:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlTiposDeClientes" runat="server" Width="392px" Enabled="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Cta.Orçamentária:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlContasOrcamentarias" runat="server" Width="392px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Responsabilidades:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlResponsabilidadeDeContas" runat="server" Width="392px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Dacon:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlDacon" runat="server" Width="392px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Tipo de Custo:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlTipoDeCusto" runat="server" Width="392px">
                                        <asp:ListItem>NENHUM</asp:ListItem>
                                        <asp:ListItem>FIXO</asp:ListItem>
                                        <asp:ListItem>VARIAVEL</asp:ListItem>
                                    </asp:DropDownList>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Ativo:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlAtivo" runat="server" Width="150px">
                                        <asp:ListItem Value="1">Ativo</asp:ListItem>
                                        <asp:ListItem Value="0">Inativo</asp:ListItem>
                                    </asp:DropDownList>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    No Receber:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlContasAReceber" runat="server" Width="150px" Enabled="False">
                                        <asp:ListItem Value=" "></asp:ListItem>
                                        <asp:ListItem Value="D">Debito</asp:ListItem>
                                        <asp:ListItem Value="C">Credito</asp:ListItem>
                                    </asp:DropDownList>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    No Pagar:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlContasAPagar" runat="server" Width="150px" Enabled="False">
                                        <asp:ListItem Value=" "></asp:ListItem>
                                        <asp:ListItem Value="D">Debito</asp:ListItem>
                                        <asp:ListItem Value="C">Credito</asp:ListItem>
                                    </asp:DropDownList>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Parametros:
                                </div>
                                <div class="coltxt">
                                    <asp:CheckBox ID="chkEncargo" runat="server" AutoPostBack="True" OnCheckedChanged="chkEncargo_CheckedChanged"
                                        Text="Participa no Financeiro" data-ToolTip="default" ToolTip="Selecionar os itens conforme configuração da conta." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt" style="margin-left: 125px;">
                                    <asp:CheckBox ID="chkCliente" runat="server" AutoPostBack="True" OnCheckedChanged="chkCliente_CheckedChanged"
                                        Text="Tem Cliente" data-ToolTip="default" ToolTip="Selecionar os itens conforme configuração da conta." />
                                    <asp:CheckBox ID="chkCentroDeCusto" runat="server" Text="Tem Centro de Custo"
                                        data-ToolTip="default" ToolTip="Selecionar os itens conforme configuração da conta." />
                                    <asp:CheckBox ID="chkProduto" runat="server" Text=" Tem Produto"
                                        data-ToolTip="default" ToolTip="Selecionar os itens conforme configuração da conta." />
                                    <asp:CheckBox ID="chkPedido" runat="server" Text="Tem Pedido"
                                        data-ToolTip="default" ToolTip="Selecionar os itens conforme configuração da conta." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt" style="margin-left: 125px;">
                                    <asp:CheckBox ID="chkAdiantamento" runat="server" AutoPostBack="True" OnCheckedChanged="chkAdiantamento_CheckedChanged"
                                        Text="É Adiantamento" data-ToolTip="default" ToolTip="Selecionar os itens conforme configuração da conta." />
                                    <asp:CheckBox ID="chkAdSoContabil" runat="server" Enabled="False" Text="O Adiantamento é Somente Contabil"
                                        data-ToolTip="default" ToolTip="Selecionar os itens conforme configuração da conta." />
                                    <asp:CheckBox ID="chkProdutoParaCusto" runat="server" Text="Produto para Custo"
                                        data-ToolTip="default" ToolTip="Selecionar os itens conforme configuração da conta." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Formato Saida:
                                </div>
                                <div class="coltxt">
                                    <asp:CheckBox ID="chkDetalhado" runat="server" Text="Detalhado"
                                        data-ToolTip="default" ToolTip="Selecionar para detalhar a conta." />
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabConsulta" runat="server" HeaderText="Consulta">
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="gridPlanoDeContas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridPlanoDeContas_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <ItemStyle Width="30px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Conta" HeaderText="Conta">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Titulo" HeaderText="Titulo" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Tem Cliente">
                                        <ItemTemplate>
                                            <asp:Label ID="lblTemCliente" runat="server" Text='<%# eval("TemCliente") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Tem C.Custo">
                                        <ItemTemplate>
                                            <asp:Label ID="lblTemCCusto" runat="server" Text='<%# eval("TemCentroDeCusto") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Tem Produto">
                                        <ItemTemplate>
                                            <asp:Label ID="lblTemProduto" runat="server" Text='<%# eval("TemProduto") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Tem Pedido">
                                        <ItemTemplate>
                                            <asp:Label ID="lblTemPedido" runat="server" Text='<%# eval("TemPedido") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="&#201; Adto.">
                                        <ItemTemplate>
                                            <asp:Label ID="lblTemAdto" runat="server" Text='<%# eval("Adiantamento") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Ad. S&#243; Contabil">
                                        <ItemTemplate>
                                            <asp:Label ID="lblAdSoContabil" runat="server" Text='<%# eval("AdiantamentoSoContabil") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="&#201; Encargo">
                                        <ItemTemplate>
                                            <asp:Label ID="lblTemEncargo" runat="server" Text='<%# eval("Encargo") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Está Ativo">
                                        <ItemTemplate>
                                            <asp:Label ID="lblAtivo" runat="server" Text='<%# eval("Ativo") %>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="NrEncargos" HeaderText="Nr. Enc." />
                                    <asp:BoundField DataField="ContaOrcamentaria" HeaderText="Cta.Or&#231;ament&#225;ria">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Responsabilidade" HeaderText="Resp.">
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Dacon" HeaderText="Dacon">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="TipoDeCusto" HeaderText="Tipo Custo" />
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabBacen" runat="server" HeaderText="Bacen">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoBacen" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizarBacen" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirBacen" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparBacen" Text="Limpar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conta:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtContaBacen" runat="server" MaxLength="9" CssClass="txtNumerico"
                                    Width="80px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTituloBacen" runat="server" MaxLength="80" Width="397px" Enabled="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupoBacen" runat="server" Width="500px" OnSelectedIndexChanged="ddlGrupoBacen_SelectedIndexChanged"
                                    AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produtos:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProdutoBacen" runat="server" Width="500px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Ref.Bacen:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlRefBacen" runat="server" style="width:auto;" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                C.Custos ECF:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCustoECFBacen" runat="server" Width="500px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                C.Custos ECD:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCustoECDBacen" runat="server" Width="500px" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridContaBacen" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                                        <ItemStyle Width="25px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Código">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Referencial" HeaderText="Referencial">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoGrupo" HeaderText="Grupo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoProduto" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigodeCustoECf" HeaderText="C. Custo ECF">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigodeCustoECD" HeaderText="C. Custo ECD">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
</asp:Content>
