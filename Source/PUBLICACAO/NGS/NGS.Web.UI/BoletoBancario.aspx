<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="BoletoBancario.aspx.vb" Inherits="NGS.Web.UI.BoletoBancario" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngBoletoBancarioNotaFiscal" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlBoletoBancarioNotaFiscal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Boleto Bancário Nota Fiscal
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Seleção de Dados
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeNegocio_SelectedIndexChanged"
                                    Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                                <asp:TextBox ID="txtClientes" runat="server" Width="557px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdBuscaCliente" OnClick="cmdBuscaCliente_Click" runat="server" Text=">"
                                    CssClass="btn" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default"
                                    ToolTip="Selecionar o cliente desejado." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Período:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data inicial do lançamento." />
                                &nbsp;à
                                <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data final do lançamento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Nota:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNotaFiscal" runat="server" data-ToolTip="default" ToolTip="Número da nota fiscal." />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
                    <HeaderTemplate>
                        Nota Fiscal
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkConfirmarVencimentos" runat="server" Text="Confirmar" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridBoletoNotaFiscal" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:TemplateField>
                                        <HeaderTemplate>
                                            <asp:CheckBox ID="chkTodosNotaFiscal" ToolTip="Selecionar todos" runat="server" AutoPostBack="True" OnCheckedChanged="chkTodosNotaFiscal_CheckedChanged" />
                                        </HeaderTemplate>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkNotaFiscal" runat="server" AutoPostBack="True" OnCheckedChanged="chkNotaFiscal_CheckedChanged" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="20px" />
                                        <ItemStyle Width="20px" />
                                    </asp:TemplateField>
                                    <asp:BoundField HeaderText="Empresa" DataField="Empresa">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Cliente" DataField="Cliente">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Nome" DataField="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="E/S" DataField="E/S">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="S&#233;rie" DataField="Serie">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Nota" DataField="Nota">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <asp:HiddenField ID="hidBoletoNotaFiscal" runat="server" />
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel3" runat="server" HeaderText="TabPanel3">
                    <HeaderTemplate>
                        Vencimentos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" Visible="False" ID="lnkNovo" runat="server" Text="Gravar"
                                            OnClientClick="if(!confirm('Deseja realmente agrupar estes títulos?')) return false;" OnClick="lnkNovo_Click" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparAgrupamento" OnClick="lnkLimparAgrupamento_Click" Visible="False" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkGerarRemessa" runat="server" Text="Gerar Arquivo de Remessa" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkAgruparPagamento" Text="Agrupar Pagamento" runat="server"
                                            data-tooltip="default" ToolTip="Agrupar Pagamento referente os Registros Selecionados." />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div id="divArquivo" runat="server">
                            <div class="row">
                                <div class="collbl">
                                    Banco:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlBanco" runat="server" Width="596px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Arquivo:
                                </div>
                                <div class="coltxt">
                                    <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                                        <ContentTemplate>
                                            <asp:Button ID="btnDownload" runat="server" data-ToolTip="default" ToolTip="Baixar Arquivo"
                                                Text="Download" Visible="false" />
                                        </ContentTemplate>
                                        <Triggers>
                                            <asp:PostBackTrigger ControlID="btnDownload" />
                                        </Triggers>
                                    </asp:UpdatePanel>
                                </div>
                            </div>
                        </div>
                        <div id="divAgrupamento" visible="False" runat="server" class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlAgrupamento" AutoPostBack="True" OnSelectedIndexChanged="ddlAgrupamento_SelectedIndexChanged" runat="server" Width="596px" />
                            </div>
                        </div>
                        <div id="divVencimentoAgrupamento" visible="False" runat="server" class="row">
                            <div class="collbl">
                                Vencimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVencimentoAgrupamento" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data inicial do lançamento." />
                            </div>
                            <div class="collbl">
                                Valor:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtValorAgrupamento" Enabled="False" runat="server" data-ToolTip="default" ToolTip="Número da nota fiscal." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridVencimentoBoleto" CssClass="gridSort" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" Wrap="True" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField ShowHeader="False">
                                        <HeaderTemplate>
                                            <asp:CheckBox ID="chkTodosVencimento" ToolTip="Selecionar todos" runat="server" AutoPostBack="True" OnCheckedChanged="chkTodosVencimento_CheckedChanged" />
                                        </HeaderTemplate>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkVencimento" AutoPostBack="True" OnCheckedChanged="chkVencimento_CheckedChanged" runat="server" />
                                        </ItemTemplate>
                                        <ControlStyle Width="30px" />
                                        <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                        <ItemStyle HorizontalAlign="Center" Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Registro">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:Label ID="lblHistorico" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "Historico").ToString().Replace(Environment.NewLine, "<br/>").Replace("\n", "<br />") %>'></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Prorrogacao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorLiquido" DataFormatString="{0:n2}" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField ShowHeader="False">
                                        <ItemTemplate>
                                            <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                                                <ContentTemplate>
                                                    <asp:ImageButton ID="imgBoletoPDFAntesDoEnvio" Visible="False" runat="server" ImageUrl="~/images/icopdf16X16.jpg"
                                                        ImageAlign="AbsMiddle" OnCommand="btnBoleto_Click" CommandName='<%# Eval("Codigo") %>' data-ToolTip="default"
                                                        ToolTip="Boleto Bancário" />
                                                </ContentTemplate>
                                                <Triggers>
                                                    <asp:PostBackTrigger ControlID="imgBoletoPDFAntesDoEnvio" />
                                                </Triggers>
                                            </asp:UpdatePanel>
                                        </ItemTemplate>
                                        <ControlStyle Width="60px" />
                                        <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                        <ItemStyle HorizontalAlign="Center" Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel4" runat="server">
                    <HeaderTemplate>
                        Processa Retorno Bancário
                    </HeaderTemplate>
                    <ContentTemplate>
                        <asp:UpdatePanel ID="upRetornoBancario" runat="server">
                            <ContentTemplate>
                                <div class="menu_acoes">
                                    <div class="acoes">
                                        <ul>
                                            <li runat="server">
                                                <asp:LinkButton class="iconRelatorio" ID="lnkProcessarRetorno" runat="server" Text="Processar" />
                                            </li>
                                            <li runat="server">
                                                <asp:LinkButton class="iconConsultar" ID="lnkConsultarRetorno" Text="Consultar" runat="server" />
                                            </li>
                                        </ul>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Convênio:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlBancoRetorno" runat="server" Width="596px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Arquivo:
                                    </div>
                                    <div class="coltxt">
                                        <asp:FileUpload ID="fup" ClientIDMode="Static" runat="server" Width="575px" onchange="this.form.submit();" />
                                        <asp:Button ID="btnProcessarRetorno" runat="server" Style="display: none" />
                                        <div id="divRetorno" runat="server" visible="false">
                                            <asp:Label ID="lblArquivoRetorno" runat="server"></asp:Label>
                                            <asp:ImageButton ID="imbExcluirRetorno" OnClick="imbExcluirRetorno_Click" ToolTip="Remover arquivo" ImageUrl="~/Images/borracha.jpg" runat="server" />
                                        </div>
                                    </div>
                                </div>
                                <div class="bordagrid">
                                    <asp:GridView ID="gridRetornoTitulos" runat="server" AutoGenerateColumns="False"
                                        CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <EditRowStyle BackColor="#999999" />
                                        <Columns>
                                            <asp:TemplateField ShowHeader="False">
                                                <ItemTemplate>
                                                    <asp:CheckBox ID="chkVencimento" Enabled="False" runat="server" />
                                                </ItemTemplate>
                                                <ControlStyle Width="30px" />
                                                <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                                <ItemStyle HorizontalAlign="Center" Width="30px" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="Registro">
                                                <ItemTemplate>
                                                    <asp:HyperLink style="display:none;" ID="hpTitulo" runat="server"  NavigateUrl="#" Text='<%# Eval("Codigo")%>' />
                                                    <a href="#" onclick="newTab('<%# Eval("CodigoCifrado", "WFTitulo.aspx?param={0}")%>')"><%# Eval("Codigo")%></a>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="Cidade" HeaderText="Empresa">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:TemplateField>
                                                <ItemTemplate>
                                                    <asp:Label ID="lblHistorico" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "Historico").ToString().Replace(Environment.NewLine, "<br/>").Replace("\n", "<br />") %>'></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="Observacoes" HeaderText="Situação">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Prorrogacao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ValorLiquido" DataFormatString="{0:n2}" HeaderText="Valor">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Juros" DataFormatString="{0:n2}" HeaderText="Juros">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:TemplateField ShowHeader="False">
                                                <ItemTemplate>
                                                    <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                                                        <ContentTemplate>
                                                            <asp:ImageButton ID="imgBoletoPDF" runat="server" ImageUrl="~/images/icopdf16X16.jpg"
                                                                ImageAlign="AbsMiddle" OnCommand="btnBoletoRetorno_Click" CommandName='<%# Eval("Codigo") %>' data-ToolTip="default"
                                                                ToolTip="Impressão do Boleto Bancário" />
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:PostBackTrigger ControlID="imgBoletoPDF" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </ItemTemplate>
                                                <ControlStyle Width="60px" />
                                                <HeaderStyle HorizontalAlign="Center" Width="60" />
                                                <ItemStyle HorizontalAlign="Center" Width="60" />
                                            </asp:TemplateField>
                                            <asp:TemplateField ShowHeader="False">
                                                <HeaderTemplate>
                                                    <asp:UpdatePanel ID="updpnlArquivos" runat="server">
                                                        <ContentTemplate>
                                                            <asp:ImageButton ID="imgBoletos" runat="server" ImageUrl="~/images/icopdf16X16.jpg"
                                                                ImageAlign="AbsMiddle" OnCommand="btnBoletos_Click" data-ToolTip="default"
                                                                ToolTip="Impressão de Boletos Bancários" />
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:PostBackTrigger ControlID="imgBoletos" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:CheckBox ID="chkImpBoletos" runat="server" />
                                                </ItemTemplate>
                                                <ControlStyle Width="60px" />
                                                <HeaderStyle HorizontalAlign="Center" Width="60px" />
                                                <ItemStyle HorizontalAlign="Center" Width="60px" />
                                            </asp:TemplateField>
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
