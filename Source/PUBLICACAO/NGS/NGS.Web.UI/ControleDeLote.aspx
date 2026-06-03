<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ControleDeLote.aspx.vb" Inherits="NGS.Web.UI.ControleDeLote" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript" language="javascript">
        function pageLoadControleLote() {
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
            pageLoadControleLote();
            var prmControleLote = Sys.WebForms.PageRequestManager.getInstance();
            prmControleLote.add_endRequest(pageLoadControleLote);
        });
    </script>

    <style type="text/css">
        .collbl {
            width: 135px;
        }
    </style>

</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngControleDeLote" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating1" runat="server" Text="Carregando..." />
    <asp:UpdatePanel ID="updpnlControleDeLote" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                <label>
                    Controle de Lote
                </label>
            </div>

            <ajaxToolkit:TabContainer ID="tcControleDeLote" runat="server" ActiveTabIndex="0"
                Width="100%">
                <ajaxToolkit:TabPanel runat="server" ID="tbLote" HeaderText="Lote">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" /></li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparConsulta" Text="Limpar" runat="server" /></li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" /></li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Fornecedor: 
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeFornecedor" runat="server" Enabled="False" Width="470px" />
                                <asp:Button ID="btnFornecedor" runat="server" OnClick="btnFornecedor_Click" Text=" &gt; " class="btn" data-ToolTip="default" ToolTip="Selecionar o fornecedor desejado." />
                                <asp:HiddenField ID="txtCodigoFornecedor" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Safra: 
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSafraConsulta" runat="server" Width="500px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tipo: 
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlConsultaTipoDeLote" runat="server" Width="500px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt" style="width: 100%; margin-bottom: 4px;">
                                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                            </div>
                        </div>

                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    Tipo de Consulta: 
                                </div>
                                <div class="coltxt">
                                    <asp:RadioButton ID="RdLote" runat="server" Checked="True" GroupName="Consulta" Text="Por Lote"
                                        data-ToolTip="default" ToolTip="Selecionar se a consulta é por lote ou lote e classificação." />
                                    <asp:RadioButton ID="RdLoteClassificacao" runat="server" GroupName="consulta" Text="Por Lote e Classificação"
                                        data-ToolTip="default" ToolTip="Selecionar se a consulta é por lote ou lote e classificação." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Situação: 
                                </div>
                                <div class="coltxt">
                                    <asp:RadioButton ID="RdAbertos" runat="server" GroupName="situacao" Text="Abertos"
                                        data-ToolTip="default" ToolTip="Selecionar entre abertos e vencidos." />
                                    <asp:RadioButton ID="RdVencidos" runat="server" GroupName="situacao" Text="Vencidos"
                                        data-ToolTip="default" ToolTip="Selecionar entre abertos e vencidos." />
                                    <asp:RadioButton ID="RdTodosAbertosVencidos" runat="server" Checked="True" GroupName="situacao" Text="Todos"
                                        data-ToolTip="default" ToolTip="Selecionar entre abertos e vencidos." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Saldo: 
                                </div>
                                <div class="coltxt">
                                    <asp:RadioButton ID="RdSemSaldo" runat="server" GroupName="SD" Text="Sem Saldo"
                                        data-ToolTip="default" ToolTip="Marcar para exibir com ou sem saldo." />
                                    <asp:RadioButton ID="RdComSaldo" runat="server" GroupName="SD" Text="Com Saldo"
                                        data-ToolTip="default" ToolTip="Marcar para exibir com ou sem saldo." />
                                    <asp:RadioButton ID="RdTodosSaldo" runat="server" Checked="True" GroupName="SD" Text="Todos"
                                        data-ToolTip="default" ToolTip="Marcar para exibir com ou sem saldo." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Posição dia: 
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPosicaoDia" runat="server" CssClass="calendario" Width="78px"
                                        data-ToolTip="default" ToolTip="Selecionar o dia para verificar a posição do lote." />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    Lote: 
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtLoteConsulta" runat="server" data-ToolTip="default"
                                        ToolTip="Informar o número do lote." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    <asp:CheckBox ID="chkLancados" runat="server" Text="Lançados:" AutoPostBack="True" data-ToolTip="default" ToolTip="" />
                                </div>
                                <div class="coltxt">
                                    <asp:Panel ID="PnlLancados" runat="server" Visible="False">
                                        &#160;de
                                    <asp:TextBox ID="txtDataLancamentoInicial" runat="server" CssClass="calendario" Width="70px"
                                        data-ToolTip="default" ToolTip="Informar o período dos lotes lançados para consulta." />
                                        &#160;a
                                    <asp:TextBox ID="txtDataLancamentoFinal" runat="server" CssClass="calendario" Width="70px"
                                        data-ToolTip="default" ToolTip="Informar o período dos lotes lançados para consulta." />
                                    </asp:Panel>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    <asp:CheckBox ID="chkValidade" runat="server" AutoPostBack="True" Text="Validade:" data-ToolTip="default" ToolTip="" />
                                </div>
                                <div class="coltxt">
                                    <asp:Panel ID="pnlValidade" runat="server" Visible="False">
                                        de
                                    <asp:TextBox ID="txtDataVencimentoInicial" runat="server" CssClass="calendario" Width="70px"
                                        data-ToolTip="default" ToolTip="Informar a validade dos lotes para consulta." />
                                        &#160; a&#160;
                                    <asp:TextBox ID="txtDataVencimentoFinal" runat="server" CssClass="calendario" Width="70px"
                                        data-ToolTip="default" ToolTip="Informar a validade dos lotes para consulta." />
                                    </asp:Panel>
                                </div>
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridLote" runat="server" AllowSorting="True" AutoGenerateColumns="False"
                                CellPadding="4" EnableTheming="False" GridLines="Horizontal" OnSelectedIndexChanged="gridLote_SelectedIndexChanged"
                                Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:BoundField DataField="Produto_Id" HeaderText="Produto" />
                                    <asp:BoundField DataField="Nome" HeaderText="Nome" />
                                    <asp:BoundField DataField="Lote" HeaderText="Lote" HtmlEncode="False" HtmlEncodeFormatString="False" />
                                    <asp:BoundField DataField="Classificacao" HeaderText="Classificacao" HtmlEncode="False"
                                        HtmlEncodeFormatString="False" />
                                    <asp:BoundField DataField="Cliente_Id" HeaderText="Fabricante" HtmlEncode="False"
                                        HtmlEncodeFormatString="False" />
                                    <asp:BoundField DataField="NomeFabricante" HeaderText="Nome" />
                                    <asp:BoundField DataField="Cidade" HeaderText="Cidade" />
                                    <asp:BoundField DataField="Estado" HeaderText="Estado" />
                                    <asp:BoundField DataField="Saldo" HeaderText="Saldo">
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoSafra" HeaderText="Safra" />
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="tbMovimentacao" HeaderText="Movimentação">
                    <ContentTemplate>
                        <div class="bordagrid" style="height: 100%">
                            <asp:GridView ID="gridMovimentacao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" Height="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Doc" HeaderText="Doc."></asp:BoundField>
                                    <asp:BoundField DataField="EntradaSaida_id" HeaderText="E/S"></asp:BoundField>
                                    <asp:BoundField DataField="Movimento" HeaderText="Data"></asp:BoundField>
                                    <asp:BoundField DataField="NomeEmpresa" HeaderText="Empresa"></asp:BoundField>
                                    <asp:BoundField DataField="Cidade" HeaderText="Cidade"></asp:BoundField>
                                    <asp:BoundField DataField="EstadoEmpresa" HeaderText="Estado"></asp:BoundField>
                                    <asp:BoundField DataField="Nome" HeaderText="Cliente"></asp:BoundField>
                                    <asp:BoundField DataField="Cidade" HeaderText="Cidade"></asp:BoundField>
                                    <asp:BoundField DataField="Estado" HeaderText="Estado"></asp:BoundField>
                                    <asp:BoundField DataField="Nota_Id" HeaderText="Nota"></asp:BoundField>
                                    <asp:BoundField DataField="Lote" HeaderText="Lote"></asp:BoundField>
                                    <asp:BoundField DataField="Classificacao" HeaderText="Classificacao"></asp:BoundField>
                                    <asp:BoundField DataField="QuantidadeFiscal" HeaderText="Qtde">
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="tbManutencao" HeaderText="Manutenção">
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

                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tipo: 
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTipoDeLote" runat="server" Width="500px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Lote: 
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtLote" runat="server" data-ToolTip="default" ToolTip="Informar o número do lote." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto: 
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodProdutoCadastro" runat="server" Enabled="False" Width="100px" />
                                <asp:TextBox ID="txtNomeProdutoCad" runat="server" Enabled="False" Width="320px" />
                                <asp:Button ID="btnProdutoCad" runat="server" OnClick="btnProdutoCad_Click" Text=" &gt; " class="btn"
                                    data-ToolTip="default" ToolTip="Forma de compra/venda do produto." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Safra: 
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSafra" runat="server" Width="430px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data de Lançamento: 
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataLancamento" runat="server" CssClass="calendario"
                                    data-ToolTip="default" ToolTip="Inserir o período de lançamento." />
                            </div>
                            <div class="collbl">
                                Data de Validade: 
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataValidade" runat="server" CssClass="calendario"
                                    data-ToolTip="default" ToolTip="Inserir a data de validade." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Fornecedor: 
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoFornecedorCad" runat="server" />
                                <asp:TextBox ID="txtFornecedorCad" runat="server" Enabled="False" Width="430px" />
                                <asp:Button ID="btnFornecedorCad" runat="server" OnClick="btnFornecedorCad_Click1" Text=" &gt; " class="btn"
                                    data-ToolTip="default" ToolTip="Informar o forncedor desejado." />
                            </div>
                        </div>
                        <div class="row">
                            <asp:Panel ID="pnlCamposLoteSemente" runat="server">
                                <div class="painelleft">
                                    <div class="row">
                                        <div class="collbl">
                                            Renasem: 
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtRenasem" runat="server" data-ToolTip="default"
                                                ToolTip="Número de inscrição no Renasem." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Termo de Conformidade: 
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtTermConformidade" runat="server" data-ToolTip="default"
                                                ToolTip="Número de registro do termo de conformidade." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Boletim: 
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtBoletim" runat="server" data-ToolTip="default"
                                                ToolTip="Número do boletim." />
                                        </div>
                                    </div>
                                </div>
                                <div class="painelleft">

                                    <div class="row">
                                        <div class="collbl">
                                            Pureza: 
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtPureza" runat="server" data-ToolTip="default"
                                                ToolTip="Grau de pureza." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Germinação: 
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtGerminacao" runat="server" data-ToolTip="default"
                                                ToolTip="Grau de germinação." />
                                        </div>
                                    </div>
                                </div>
                            </asp:Panel>
                        </div>
                        <asp:Panel ID="pnlClassificacao" runat="server">
                            <div class="subtitulodiv">
                                Cadastro Classificação / Peneira 
                            </div>
                            <div class="menu_acoes">
                                <div class="acoes">
                                    <ul>
                                        <li runat="server">
                                            <asp:LinkButton class="iconNovo" ID="lnkAdicionarClassificacao" Text="Adicionar" runat="server" />
                                        </li>
                                        <li runat="server">
                                            <asp:LinkButton class="iconExcluir" ID="lnkRemoverClassificacao" Text="Remover" runat="server" />
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Código: 
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtClassificacao" runat="server" Width="80px" MaxLength="10"
                                        data-ToolTip="default" ToolTip="Código de classificação do produto após passar pela peneira." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Peso do Saco: 
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPesoSaco" runat="server" CssClass="txtDecimal4" Width="80px"
                                        data-ToolTip="default" ToolTip="Peso do saco após passar pelas peneiras de classificação." />
                                </div>
                            </div>
                            <div class="bordagrid">
                                <asp:GridView ID="gridClassificacao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridClassificacao_SelectedIndexChanged">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                        <asp:BoundField DataField="Classificacao" HeaderText="Classificação" />
                                        <asp:BoundField DataField="PesoSaco" DataFormatString="{0:N4}" HeaderText="Peso Saco" />
                                    </Columns>
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                </asp:GridView>
                            </div>
                        </asp:Panel>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>
