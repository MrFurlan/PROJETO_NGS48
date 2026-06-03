<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ControleDeFretes.aspx.vb" Inherits="NGS.Web.UI.ControleDeFretes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngControleDeFretes" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlControleDeFretes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Controle de Fretes
            </div>
            <ajaxToolkit:TabContainer ID="TabLancamento" runat="server" ActiveTabIndex="1" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="Pré-Fatura">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoPF" runat="server" Text="Gravar" />
                                    </li>
                                    <%--  <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkAtualizarPF" runat="server" Text="Atualizar" />
                                    </li>--%>

                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizarPF" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirPF" runat="server" Text="Excluir"
                                            OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultarPF" runat="server" Text="Consultar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatorioPF" runat="server" Text="Relatório" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparPF" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaPF" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtUnNegocio" runat="server" />
                                <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                                <asp:TextBox ID="txtEmpresa" runat="server" Enabled="False" Width="552px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnEmpresa" runat="server" UseSubmitBehavior="False" Text=">" CssClass="btn"
                                    OnClick="btnEmpresa_Click" data-ToolTip="default" ToolTip="Empresa Para negociação/consulta." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conveniado:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="CodigoConveniadoPF" runat="server" />
                                <asp:TextBox ID="txtConveniadoPF" runat="server" Width="552px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnConveniadoPF" OnClick="btnConveniadoPF_Click" CssClass="btn" runat="server"
                                    UseSubmitBehavior="False" Text=">" data-ToolTip="default" ToolTip="Transportadora conveniada." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Fatura:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtFaturaPF" runat="server" CssClass="txtNumerico" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnDadosBancarios" OnClick="btnDadosBancarios_Click" runat="server"
                                    UseSubmitBehavior="False" CssClass="botao" Width="260px" Text="Carregar Conta Cadastrada do Conveniado"
                                    data-ToolTip="default" ToolTip="Inserir o númdo de registro do CTRC ou da fatura." />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdbPagar" runat="server" Text="Contas A Pagar" Checked="True" GroupName="PR"
                                    data-ToolTip="default" ToolTip="Contas A Pagar. " />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdbReceber" runat="server" Text="Contas A Receber" GroupName="PR"
                                    data-ToolTip="default" ToolTip="Contas A  Receber. " />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Valor Fatura:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtValorFaturaPF" CssClass="txtDecimal" runat="server" data-ToolTip="default"
                                    ToolTip="Valor Total da Fatura para pagamento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Banco:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodBanco" runat="server" Width="42px" Enabled="False" data-ToolTip="default"
                                    ToolTip="Selecionar o banco de pagamento/recebimento." />
                                <asp:TextBox ID="txtNomeBanco" runat="server" Width="160px" Enabled="False" data-ToolTip="default"
                                    ToolTip="Selecionar o banco de pagamento/recebimento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Movimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMovimentoPF" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data do lançamento." />
                            </div>
                            <div class="collbl">
                                Agência:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodAgencia" runat="server" Width="42px" Enabled="False" data-ToolTip="default"
                                    ToolTip="Selecionar a agencia que será pago/recebido o valor." />
                                <asp:TextBox ID="txtDigitoAgencia" runat="server" Width="20px" Enabled="False" MaxLength="2"
                                    data-ToolTip="default" ToolTip="Selecionar a agencia que será pago/recebido o valor." />
                                <asp:TextBox ID="txtPracaAgencia" runat="server" Enabled="False" data-ToolTip="default"
                                    ToolTip="Selecionar a agencia que será pago/recebido o valor." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Vencimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVencimentoPF" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data que vence o pagamento/recebimento. " />
                            </div>
                            <div class="collbl">
                                Conta:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtConta" runat="server" Width="42px" Enabled="False" data-ToolTip="default"
                                    ToolTip="Selecionar a conta que será pago/recebido o valor." />
                                <asp:TextBox ID="txtDigitoConta" runat="server" Width="20px" Enabled="False" MaxLength="2"
                                    data-ToolTip="default" ToolTip="Selecionar a conta que será pago/recebido o valor." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Carteira:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCarteiraPF" runat="server" Width="583px" data-ToolTip="default"
                                    ToolTip="Informar a transação que será realizada." />
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            Pré Faturas de Frete
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Vencimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtConsultaDePF" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data que vence o pagamento/recebimento. " />
                            </div>
                            <div class="collbl">
                                Até:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtConsultaAtePF" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data que vence o pagamento/recebimento" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdAtivoPF" runat="server" Text="Ativos" Checked="True" GroupName="PF"
                                    data-ToolTip="default" ToolTip="Data que vence o pagamento/recebimento. " />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdBaixadosPF" runat="server" Text="Baixados" GroupName="PF"
                                    data-ToolTip="default" ToolTip="Data que vence o pagamento/recebimento. " />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="DgPreFaturas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="DgPreFaturas_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                                        <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:CommandField>

                                    <asp:TemplateField HeaderText="Empresa">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <%#Eval("Titulo.CodigoEmpresa")%>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="End.">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                        <ItemTemplate>
                                            <%#Eval("Titulo.EnderecoEmpresa")%>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoFatura" HeaderText="Fatura" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoConveniado" HeaderText="Conveniado" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EnderecoConveniado" HeaderText="End" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ConveniadoNome" HeaderText="Nome" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorDaFatura" DataFormatString="{0:N2}" HeaderText="Valor"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Titulo">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                        <ItemTemplate>
                                            <asp:Label ID="lblTitulo" runat="server" Text='<%#Eval("Titulo.Codigo")%>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Vencimento">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" />
                                        <ItemTemplate>
                                            <asp:Label ID="lblTitulo" runat="server" Text='<%#CDate(Eval("Titulo.Prorrogacao")).ToString("dd/MM/yyyy")%>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="Lançamentos">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatório" runat="server" Text="Relatório" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <asp:Panel ID="pnlMsg" runat="server" Visible="False">
                            <div class="ui-widget">
                                <div class="ui-state-highlight ui-corner-all" style="padding: 0 .7em;">
                                    <p>
                                        <span class="ui-icon ui-icon-info" style="float: left; margin-right: .3em;"></span>
                                        <strong>Atenção!</strong>
                                        <asp:Label ID="lblMsg" runat="server" />
                                    </p>
                                </div>
                            </div>
                        </asp:Panel>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Width="656px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conveniado:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="CodigoConveniadoLan" runat="server" />
                                <asp:TextBox ID="txtConveniadoLan" runat="server" Enabled="False" Width="617px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnConveniadoLan" runat="server" UseSubmitBehavior="False" Text=">"
                                    CssClass="btn" OnClick="btnConveniado_Click" data-ToolTip="default" ToolTip="Transportadora conveniada." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Fatura/CTRC:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="CodigoFatura" runat="server" />
                                <asp:TextBox ID="txtFatura" CssClass="txtNumerico9" runat="server" data-ToolTip="default"
                                    ToolTip="Inserir o númdo de registro do CTRC ou da fatura." />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton Style="cursor: pointer" ID="imgFatura" OnClick="imgFatura_Click"
                                    CssClass="imgconsultar" runat="server" Enabled="False" data-ToolTip="default"
                                    ToolTip="Consultar Fatura" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                    ImageAlign="AbsMiddle"></asp:ImageButton>
                            </div>
                            <div class="collbl">
                                Movimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataMovimento" CssClass="calendario" runat="server" Enabled="False"
                                    data-ToolTip="default" ToolTip="Data do lançamento." />
                            </div>
                            <div class="collbl">
                                Vencimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataVencimento" CssClass="calendario" runat="server" Enabled="False"
                                    data-ToolTip="default" ToolTip="Data que vence o pagamento. " />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Valor Fatura:
                            </div>
                            <div class="coltxt" style="width: 130px;">
                                <asp:TextBox ID="txtValorFatura" CssClass="txtDecimal" runat="server" ReadOnly="True"
                                    Enabled="False" data-ToolTip="default" ToolTip="Valor Total da Fatura para pagamento." />
                            </div>
                            <div class="collbl">
                                Valor Lançado:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtValorLancado" CssClass="txtDecimal" runat="server" ReadOnly="True"
                                    Enabled="False" />
                            </div>
                            <div class="collbl">
                                Valor de Saldo:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtValorSaldo" CssClass="txtDecimal" runat="server" ReadOnly="True"
                                    Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdbPagarLancamento" runat="server" Text="Pagar" Checked="True" GroupName="PRL"
                                    data-ToolTip="default" ToolTip="Contas A Pagar. " />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdbReceberLancamento" runat="server" Text="Receber" GroupName="PRL"
                                    data-ToolTip="default" ToolTip="Contas A  Receber. " />
                            </div>
                        </div>
                        <ajaxToolkit:TabContainer ID="TabOutros" runat="server" Width="100%" ActiveTabIndex="1">
                            <ajaxToolkit:TabPanel ID="TabPanel4" runat="server" HeaderText="Dados do Frete">
                                <ContentTemplate>
                                    <div class="row">
                                        <div class="collbl">
                                            Documento:
                                        </div>
                                        <div class="coltxt">
                                            <asp:RadioButton ID="RbCtrc" runat="server" Text="CTRC" GroupName="TipoDoc" Checked="True" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:RadioButton ID="RbRecibo" runat="server" Text="Recibo" Enabled="False" GroupName="TipoDoc" Visible="false" />
                                        </div>
                                        <div class="collbl" style="margin-left: 16px;">
                                            Via:
                                        </div>
                                        <div class="coltxt">
                                            <asp:RadioButton ID="optAdiantamento" runat="server" Text="Adiantamento" Checked="True"
                                                GroupName="Via" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:RadioButton ID="optCartaFrete" runat="server" Text="Saldo" GroupName="Via" OnCheckedChanged="optCartaFrete_CheckedChanged"
                                                AutoPostBack="true" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:RadioButton ID="optAmortizacao" runat="server" Text="Amortização" Enabled="False"
                                                GroupName="Via" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:RadioButton ID="optBaixaAdiantamento" OnCheckedChanged="optBaixaAdiantamento_CheckedChanged"
                                                AutoPostBack="true" runat="server" Text="Baixa de Adiantamento" GroupName="Via" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Número:
                                        </div>
                                        <div class="coltxt">
                                            <asp:HiddenField ID="NotasDeFrete" runat="server" />
                                            <asp:TextBox ID="TxtNrFrete" runat="server" CssClass="txtNumerico9" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:Button ID="BtnNumeroNotaFrete" runat="server" OnClick="BtnNumeroNotaFrete_Click"
                                                Text=">" CssClass="btn" UseSubmitBehavior="False" />
                                        </div>
                                        <div class="collbl">
                                            Vencimento:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtVencimento" runat="server" CssClass="calendario" data-ToolTip="default"
                                                ToolTip="Data que vence o pagamento/recebimento. " />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Filial:
                                        </div>
                                        <div class="coltxt">
                                            <asp:HiddenField ID="hdFilial" runat="server" />
                                            <asp:DropDownList ID="ddlFilial" runat="server" Width="660px" Enabled="False" data-ToolTip="default"
                                                ToolTip="Selecionar a filial que está efetuand o pagamento ou recebendo." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Transportador:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtTransportador" runat="server" Width="655px" Enabled="False" data-ToolTip="default"
                                                ToolTip="Nome da pessoa/empresa responsável pelo transporte da mercadoria." />
                                        </div>
                                        <div class="coltxt">
                                            <asp:Button ID="Button1" runat="server" Text=">" CssClass="btn" Enabled="False" data-ToolTip="default"
                                                ToolTip="Nome da pessoa/empresa responsável pelo transporte da mercadoria." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Peso:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="TxtPeso" CssClass="txtNumerico9" runat="server" Enabled="False"
                                                data-ToolTip="default" ToolTip="Peso original da mercadoria." />
                                        </div>
                                        <div class="collbl">
                                            Adiantamento:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtPercAdiantamento" CssClass="txtNumerico9" runat="server" Width="60px"
                                                Enabled="False" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="TxtValorAdiantamento" CssClass="txtDecimal" runat="server" Enabled="False"
                                                data-ToolTip="default" ToolTip="" />
                                        </div>
                                        <div class="collbl">
                                            Peso Chegada:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtPesoChegada" CssClass="txtNumerico9" runat="server" OnTextChanged="txtPesoChegada_TextChanged"
                                                AutoPostBack="True" data-ToolTip="default" ToolTip="Peso quando da mercadoria na chegada do produto." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Frete Combinado:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtFreteCombinado" CssClass="txtDecimal" runat="server" Enabled="False" />
                                        </div>
                                        <div class="collbl">
                                            Valor KG:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtValorKg" CssClass="txtDecimal" runat="server" Enabled="False" />
                                        </div>
                                        <div class="collbl" style="margin-left: 74px;">
                                            Frete Chegada:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtFreteChegada" CssClass="txtDecimal" runat="server" Enabled="False" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Valor Frete:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtValorFrete" CssClass="txtDecimal" runat="server" Enabled="False" />
                                        </div>
                                        <div class="collbl">
                                            Tolerância KG:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtTolerancia" CssClass="txtNumerico9" runat="server" Enabled="False"
                                                data-ToolTip="default" ToolTip="Qual a margem de tolerância para perdas do produto em kg." />
                                        </div>
                                        <div class="collbl" style="margin-left: 74px;">
                                            Quebra/Sobra(R$):
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtDiferenca" CssClass="txtDecimal" runat="server" Enabled="False"
                                                data-ToolTip="default" ToolTip="Valor em reais das sobras ou quebras de produto." />
                                        </div>
                                    </div>

                                    <div class="row">
                                        <div class="collbl">
                                            Valor. Cte Comp.:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtValorCteComplementar" CssClass="txtDecimal" runat="server" Enabled="False" />
                                        </div>

                                        <div class="collbl">
                                            Valor Total do Frete:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtValorTotalFrete" CssClass="txtDecimal" runat="server" Enabled="False" />
                                        </div>


                                    </div>

                                    <div class="bordagrid" style="height: 250px;">
                                        <asp:GridView ID="dgEncargos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                            ForeColor="#333333" GridLines="None" Width="100%">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EditRowStyle BackColor="#999999" />
                                            <Columns>
                                                <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:CommandField>
                                                <asp:BoundField DataField="Codigo" HeaderText="Encargo">
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="ContaDeDebito" HeaderText="Debitar">
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="ContaDeCredito" HeaderText="Creditar">
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:BoundField>
                                                <asp:TemplateField HeaderText="Valor">
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtValorEncargo" runat="server" CssClass="txtDecimal" Text='<%# Eval("Valor", "{0:N2}") %>'
                                                            Enabled="False" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="Sinal" HeaderText="Sinal">
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:BoundField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabPanel5" runat="server" HeaderText="Fretes que Compõem a Fatura">
                                <ContentTemplate>
                                    <div class="bordagrid" style="height: 415px;">
                                        <asp:GridView ID="DgComposicaoFatura" runat="server" AutoGenerateColumns="False"
                                            CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="DgComposicaoFatura_SelectedIndexChanged">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EditRowStyle BackColor="#999999" />
                                            <Columns>
                                                <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:CommandField>
                                                <asp:BoundField DataField="CodigoEncargo" HeaderText="Encargo" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:BoundField>
                                                <asp:TemplateField HeaderText="Recibo">
                                                    <ItemTemplate>
                                                        <asp:CheckBox ID="ckReciboDeFrete" runat="server" Checked='<%# Eval("ReciboFrete") %>'
                                                            Enabled="False" />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="CTRC">
                                                    <ItemTemplate>
                                                        <asp:CheckBox ID="ckCTRC" runat="server" Checked='<%# Eval("CTRC") %>' Enabled="False" />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="NumeroNota" HeaderText="Nota" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Peso" HeaderText="Peso" DataFormatString="{0:n0}">
                                                    <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="ValorLancadoNota" HeaderText="Valor Nota" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="EmpresaCnpj" HeaderText="Empresa" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:BoundField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabVencimentosNovo" runat="server" HeaderText="Vencimentos">
                                <HeaderTemplate>
                                    Vencimentos
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <uc:Financeiro ID="ucFinanceiro" runat="server" />
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                        </ajaxToolkit:TabContainer>
                        <div class="row">
                            <asp:Button ID="btnOk" runat="server" CssClass="botao" UseSubmitBehavior="False"
                                Text="OK" OnClick="btnOk_Click"></asp:Button>
                            <asp:Button ID="btnExcluir" runat="server" CssClass="botao" UseSubmitBehavior="False"
                                Text="Excluir" OnClick="btnExcluir_Click"></asp:Button>
                            <asp:Button ID="btnLimparLancamento" runat="server" CssClass="botao" UseSubmitBehavior="False"
                                Text="Limpar" OnClick="btnLimparLancamento_Click"></asp:Button>
                            <asp:Button ID="btnCalcular" runat="server" CssClass="botao" Enabled="False" UseSubmitBehavior="False"
                                OnClick="btnCalcular_Click" Text="Calcular" />
                            <asp:Button ID="btnEncerrar" runat="server" CssClass="botao" Enabled="False" UseSubmitBehavior="False"
                                OnClick="btnEncerrar_Click" Text="Encerrar" />
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel3" runat="server" HeaderText="Programação">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultarPR" runat="server" Text="Consultar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparPR" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaPR" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaProg" runat="server" Width="575px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conveniado:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="CodigoConveniadoProg" runat="server" />
                                <asp:TextBox ID="txtConveniadoProg" runat="server" Width="543px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnConveniadoProg" OnClick="btnConveniadoProg_Click" runat="server"
                                    UseSubmitBehavior="False" Text=">" CssClass="btn" data-ToolTip="default" ToolTip="Transportadora conveniada." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="CodigoClienteProg" runat="server" />
                                <asp:TextBox ID="txtClienteProg" runat="server" Width="543px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnClienteProg" OnClick="btnClienteProg_Click" CssClass="btn" runat="server"
                                    UseSubmitBehavior="False" Text=">" data-ToolTip="default" ToolTip="Cliente CTE." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Fatura/CTRC:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNumeroFatura" CssClass="txtNumerico9" runat="server" data-ToolTip="default"
                                    ToolTip="Inserir o númdo de registro do CTRC ou da fatura." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Vencimento:
                            </div>
                            <div class="collbl">
                                <asp:CheckBox ID="chkUsarPeriodo" runat="server" AutoPostBack="True" Text="Usar Periodo"
                                    Checked="false" />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVencimentoDe" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data que vence o pagamento/recebimento. " />
                            </div>
                            <div class="collbl">
                                Até:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVencimentoAte" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data que vence o pagamento/recebimento. " />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdAtivos" runat="server" Text="Ativos" Checked="True" GroupName="Prog" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdLiberados" runat="server" Text="Liberados" GroupName="Prog" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdBaixados" runat="server" Text="Baixados" GroupName="Prog" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="DgProgFaturas" runat="server" AutoGenerateColumns="False" CssClass="bordasimples"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>

                                    <asp:TemplateField HeaderText="Empresa">
                                        <ItemTemplate>
                                            <%#Eval("Titulo.CodigoEmpresa")%>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="End.">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                        <ItemTemplate>
                                            <%#Eval("Titulo.EnderecoEmpresa")%>
                                        </ItemTemplate>
                                    </asp:TemplateField>

                                    <asp:BoundField DataField="CodigoFatura" HeaderText="Fatura">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoConveniado" HeaderText="Conveniado">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EnderecoConveniado" HeaderText="End.">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ConveniadoNome" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
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
                                    <asp:TemplateField HeaderText="Titulo">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                        <ItemTemplate>
                                            <asp:Label ID="lblTitulo" runat="server" Text='<%#Eval("Titulo.Codigo")%>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>

                                    <asp:TemplateField HeaderText="Valor do Título">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                        <ItemTemplate>
                                            <%#Eval("Titulo.ValorDoDocumento", "{0:N2}") %>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgDelete" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/deletar.gif"
                                                Style="border: 0;" OnClick="imgDelete_Click" data-ToolTip="default" ToolTip="Bloqueia o Título novamente"
                                                OnClientClick="if(!confirm('Deseja realmente bloquear este registro?')) return false;" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaFaturasDeFrete ID="ucConsultaFaturasDeFrete" runat="server" />
    <uc:ConsultaNotasDeFrete ID="ucConsultaNotasDeFrete" runat="server" />
    <uc:ConsultaDadosBancarios ID="ucConsultaDadosBancarios" runat="server" />
</asp:Content>
