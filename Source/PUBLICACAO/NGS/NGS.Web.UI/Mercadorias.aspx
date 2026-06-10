<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Mercadorias.aspx.vb" Inherits="NGS.Web.UI.Mercadorias" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 140px !important;
        }

        .w50 {
            width: 55px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngMercadorias" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlMercadorias" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Produtos e Mercadorias
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Height="100%"
                Width="100%">
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel1" ID="TabPanel1">
                    <HeaderTemplate>
                        Cadastro
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconConsultar" runat="server">
                                        <asp:linkbutton ID="lnkConsultarPamCard" text="ConsultarPamCard" runat="server" />
                                    </li>
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
                                            <li class="iconExcel">
                                                <asp:LinkButton ID="lnkExcel" runat="server" Text="Excel" />
                                            </li>
                                            <li class="iconExcel">
                                                <asp:LinkButton ID="lnkExcelDados" runat="server" Text="Excel Dados" />
                                            </li>
                                        </ul>
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                    <div style="float: right;">
                                        <li runat="server">
                                            <div class="row" style="margin-top: 0;">
                                                <div class="coltxt">
                                                    <asp:Image ID="imgUsuario" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/man2.png"
                                                        data-ToolTip="default" ToolTip="Usuário Lançamento" Width="20px" />
                                                </div>
                                                <div class="coltxt">
                                                    <asp:DropDownList ID="ddlUsuarios" runat="server" Width="175px" />
                                                </div>
                                            </div>
                                        </li>
                                    </div>
                                </ul>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    Código Indea:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCodigoIndea" CssClass="txtNumerico" runat="server" Enabled="False" MaxLength="20"
                                        data-ToolTip="default" ToolTip="Código de registro no Instituto de Defesa Agropecuária do Estado de Mato Grosso." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Situação:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlSituacao" runat="server" Width="200px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Produto:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCodigoProduto" CssClass="txtNumerico" runat="server" Width="110px"
                                        Enabled="False" MaxLength="30" AutoPostBack="True" data-ToolTip="default" ToolTip="Código do Produto." />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton Style="cursor: pointer" ID="imgSequencia"
                                        CssClass="imgconsultar" runat="server" Enabled="False" data-ToolTip="default" ToolTip="Busca última sequência do produto."
                                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ImageAlign="AbsMiddle" />
                                </div>
                                <div id="divProdutoDeTerceiro" runat="server">
                                    <div class="collbl">
                                        Produto de Terceiro:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtCodigoProdutoTerceiro" runat="server" Width="60px" MaxLength="30" AutoPostBack="True" data-ToolTip="default" ToolTip="Código do Produto de Terceiro." />
                                    </div>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Grupo:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlGrupoProduto" runat="server" Width="366px" Font-Names="monospace" AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Nome:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNome" runat="server" Width="357px" MaxLength="100"
                                        data-ToolTip="default" ToolTip="Nome do produto." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Descrição:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDescricao" runat="server" MaxLength="355" TextMode="MultiLine"
                                        Width="358px" data-ToolTip="default" ToolTip="Descrição do produto." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl" style="color: red;">
                                    Adicional para NF-e:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtInfaDProd" runat="server" MaxLength="200" TextMode="MultiLine"
                                        Width="358px" data-ToolTip="default" ToolTip="Complemento para Descrição do produto na NFE." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Etapas:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlEtapas" runat="server" Width="208px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Qualidade:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlQualidade" runat="server" Width="208px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Agrupar:
                                </div>
                                <div class="coltxt">
                                    <asp:RadioButton ID="rdASim" runat="server" AutoPostBack="True" OnCheckedChanged="rdASim_CheckedChanged"
                                        Text="Sim" data-ToolTip="default" ToolTip="Informar se será ou não agrupado o produto." />
                                    <asp:RadioButton ID="rdANao" runat="server" AutoPostBack="True" OnCheckedChanged="rdANao_CheckedChanged"
                                        Text="Não" data-ToolTip="default" ToolTip="Informar se será ou não agrupado o produto." />
                                </div>
                                <div class="coltxt">
                                    <asp:CheckBox ID="chkControlarRomaneio" runat="server" AutoPostBack="True" Enabled="False"
                                        OnCheckedChanged="chkControlarRomaneio_CheckedChanged" Text="Controlar Romaneio" data-ToolTip="default" ToolTip="" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Descr.Mapa:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDescricaoMapa" runat="server" MaxLength="20" Width="200px" data-ToolTip="default" ToolTip="" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Código NCM:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNCM" runat="server" CssClass="txtNumerico8" Width="200px"
                                        data-ToolTip="default" ToolTip="Código que identifica a natureza das mercadorias." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    CNAE:
                                </div>
                                <div class="coltxt">
                                    <%--<asp:TextBox ID="txtCnae" runat="server" CssClass="txtNumerico7" Width="200px" data-ToolTip="default" ToolTip="Código da atividade econômica do estabelecimento (Somente Números)" MaxLength="7" />--%>
                                    <asp:DropDownList ID="ddlCnae" runat="server" Width="366px" data-ToolTip="default" ToolTip="Código da atividade econômica do estabelecimento" />

                                </div>
                            </div>
                            <div>
                                <div class="painelleft">
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkFitossanitario" runat="server" AutoPostBack="True" OnCheckedChanged="chkFitossanitario_CheckedChanged"
                                                Text="Fitossanitário" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkPrecoDePauta" runat="server" Text="Preço de Pauta" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkControlarPesagem" runat="server" Text="Controlar Pesagem" OnCheckedChanged="chkControlarPesagem_CheckedChanged"
                                                AutoPostBack="True" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkPrecoDoProduto" runat="server" Text="Preço do Produto" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkCustoIndireto" runat="server" Text="Custo indireto" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                </div>
                                <div class="painelleft">
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkPecasMeios" runat="server" Text="Controlar Peças/Meios" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkDecimais" runat="server" Text="Decimais Relatório" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkNumeroDoLote" runat="server" Text="Controlar Número do Lote" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkAutorizacaoDeRetirada" runat="server" Text="Obriga Autorização de Retirada" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkDashboard" runat="server" Text="Dashboard" data-ToolTip="default" ToolTip="Dashboard" />
                                        </div>
                                    </div>
                                </div>
                                <div class="painelleft">
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkAlmoxarifado" runat="server" Text="Controlar Almoxarifado" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkControlarEmbalagem" runat="server" Text="Controlar Embalagem"
                                                OnCheckedChanged="chkControlarEmbalagem_CheckedChanged" AutoPostBack="True" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkControlarEstoque" runat="server" Text="Controlar Estoque" data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkControlarLote" runat="server" Text="Controlar Lote" data-ToolTip="default" ToolTip="Somente para controle de sementes." />
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    Embalagem Padrão:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlEmbalagem" runat="server" Width="208px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Peso/Quantidade:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlPesoQuantidade" runat="server" Width="208px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Und. Faturam./Invent.:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlUnidadeDeMedida" runat="server" Width="208px" AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="subtitulodiv">
                                Unidade de Comercialização
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Un.Comercialização:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlUnidadedeComercializacao" runat="server" Width="208px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Peso da Embalagem:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPesoDaEmbalagem" runat="server" CssClass="txtDecimal4" Width="80px" data-ToolTip="default" ToolTip="Informar o peso da embalagem." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Peso do Produto:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPesoDoProduto" runat="server" CssClass="txtDecimal4" Width="80px"
                                        data-ToolTip="default" ToolTip="Informar o peso do produto." />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="btnAdicionarUnidComercializacao" runat="server" CssClass="btn"
                                        ImageUrl="~/images/ico-mais.gif" Width="18px" data-ToolTip="default" ToolTip="" />
                                </div>
                            </div>
                            <div class="bordagrid" style="height: 180px;">
                                <asp:GridView ID="gridUnidadeComercializacao" runat="server" AutoGenerateColumns="False"
                                    DataKeyNames="CodigoUnidade,FatorConversao"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:BoundField DataField="CodigoUnidade" HeaderText="UN">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Peso Embalagem">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="PesoDaEmbalagem" runat="server" CssClass="txtDecimal4" Text='<%# Eval("PesoDaEmbalagem", "{0:N4}") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="FatorConversao" DataFormatString="{0:N4}" HeaderText="Peso Produto">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Ativo">
                                            <ItemTemplate>
                                                <asp:Button ID="BtnAtivo" runat="server" CausesValidation="False" Font-Bold="True"
                                                    Font-Names="Arial Narrow" Font-Size="8pt" OnClick="BtnAtivo_Click" Style="cursor: pointer;"
                                                    Text='<%# bind("Ativo") %>' Width="40px" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Atualizar">
                                            <ItemTemplate>
                                                <asp:ImageButton ID="btnAtualizarUnidComercializacao" runat="server" ImageUrl="~/Images/update.gif"
                                                    OnClick="btnAtualizarUnidComercializacao_Click" OnClientClick="return confirm('Atenção! Deseja realmente atualizar o peso da embalagem desse produto?');"
                                                    Style="border: 0;" data-ToolTip="default" ToolTip="Atualizar" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Excluir">
                                            <ItemTemplate>
                                                <asp:ImageButton ID="btnExcluirUnidComercializacao" runat="server" ImageUrl="~/Images/deletar.gif"
                                                    OnClick="btnExcluirUnidComercializacao_Click" OnClientClick="return confirm('Atenção! Deseja realmente excluir esta Unidade de Comercialização?');"
                                                    Style="border: 0;" data-ToolTip="default" ToolTip="Excluir" />
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                    </Columns>
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                </asp:GridView>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    % IPI:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtIPI" runat="server" MaxLength="5" CssClass="txtDecimal" Width="200px"
                                        data-ToolTip="default" ToolTip="Informar o percentual do IPI" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Estoque mínimo:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtEstoqueMinimo" runat="server" CssClass="txtDecimal4" Width="200px" data-ToolTip="default" ToolTip="Estoque mínimo." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Estado Físico:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlEstadoFisico" runat="server" Width="208px" ToolTip="Estado Físico." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Registro Min Agr:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtRegMinAgr" runat="server" Width="200" MaxLength="25" data-ToolTip="default" ToolTip="Registro Ministrio da Agricultura." />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    Seguimento:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlSeguimentos" runat="server" Font-Size="8pt" Width="545px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Marca:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlMarca" runat="server" Font-Size="8pt" Width="545px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Carteira de Compras:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlCarteiraDeCompras" runat="server" Width="545px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Carteira Vendas:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlCarteiraDeVendas" runat="server" Width="545px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Tipo do Item:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlTipoDoItem" runat="server" Width="545px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Código do Serviço:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCodigoDoServico" CssClass="txtNumerico" runat="server" Width="200px"
                                        data-ToolTip="default" ToolTip="Informar o código de serviço anexa à Lei Complementar nº 116, de 31 de julho de 2003. http://sped.rfb.gov.br/pagina/show/1601" />
                                </div>
                                <div class="collbl">
                                    Código EX:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCodigoEX" runat="server" Width="170px" MaxLength="3"
                                        data-ToolTip="default" ToolTip="" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Código do Gênero:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlCodigoDoGenero" runat="server" Width="545px" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlCodigoDoGenero_SelectedIndexChanged" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Sub do Gênero:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlSubCodigoDoGenero" runat="server" Width="545px" />
                                </div>
                            </div>
                        </div>
                        <div class="painelright" style="width: 24%;">
                            <div class="collbl" style="padding-right: 25%;">
                                Código Gtin:
                            </div>
                            <div class="row">
                                <div class="collbl w50">
                                    GTIN-8:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtGTIN8" CssClass="txtNumerico" runat="server" Width="130px" MaxLength="8" />
                                </div>
                            </div>
                            <div class="row">
                                <asp:Image ID="imgGTIN8" runat="server" />
                            </div>
                            <div class="row">
                                <div class="collbl w50">
                                    GTIN-12:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtGTIN12" CssClass="txtNumerico" runat="server" Width="130px" MaxLength="12" />
                                </div>
                            </div>
                            <div class="row">
                                <asp:Image ID="imgGTIN12" runat="server" />
                            </div>
                            <div class="row">
                                <div class="collbl w50">
                                    GTIN-13:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtGTIN13" CssClass="txtNumerico" runat="server" Width="130px" MaxLength="13" />
                                </div>
                            </div>
                            <div class="row">
                                <asp:Image ID="imgGTIN13" runat="server" />
                            </div>
                            <div class="row">
                                <div class="collbl w50">
                                    GTIN-14:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtGTIN14" CssClass="txtNumerico" runat="server" Width="130px" MaxLength="14" />
                                </div>
                            </div>
                            <div class="row">
                                <asp:Image ID="imgGTIN14" runat="server" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel2" ID="TabEmbalagens">
                    <HeaderTemplate>
                        Embalagem
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="titulodiv">
                            Embalagens Do Produto
                        </div>
                        <div class="menu_acoes" runat="server">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkAdicionarEmbalagem" Text="Adicionar Embalagem" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    Embalagem:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlPXEmbalagem" runat="server" Width="260px" Enabled="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Material Embalagem:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlMaterialEmbalagem" runat="server" Width="260px" Enabled="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Qtde item naEmbalagem:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtQuantidadeDaEmbalagem" CssClass="txtDecimal4" runat="server"
                                        Width="57px" Enabled="False" />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="coltxt">
                                    <asp:CheckBox ID="chkPesoVariavel" runat="server" Text="Peso Variavel" AutoPostBack="True" />
                                </div>
                            </div>
                            <div id="pnlPesoEmbalagem" runat="server">
                                <div class="row">
                                    <div class="collbl">
                                        Peso Liquido:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtPesoLiquido" runat="server" Enabled="False" CssClass="txtDecimal4"
                                            Width="85px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Peso Bruto:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtPesoBruto" runat="server" Enabled="False" CssClass="txtDecimal4"
                                            Width="85px" />
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridEmbalagem" runat="server" Width="100%" ForeColor="#333333"
                                GridLines="None" CellPadding="4" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="DescricaoEmbalagem" HeaderText="Embalagem">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoTipoDeEmbalagem" HeaderText="Tipo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Capacidade" DataFormatString="{0:N4}" HeaderText="Capacidade">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="PesoBruto" DataFormatString="{0:N4}" HeaderText="Bruto">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="PesoLiquido" DataFormatString="{0:N4}" HeaderText="L&#237;quido">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:CheckBoxField DataField="PesoVariavel" HeaderText="Peso Variavel">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:CheckBoxField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton Style="cursor: pointer" ID="imgExcluirEmbalagem" OnClick="imgExcluirEmbalagem_Click"
                                                runat="server" Width="18px" Height="18px" data-ToolTip="default" ToolTip="Excluir Embalagem" ImageUrl="~/images/erro.jpg"
                                                ImageAlign="AbsMiddle"></asp:ImageButton>
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel5" ID="TabPanel5">
                    <HeaderTemplate>
                        Especificação 
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoEPrd" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizarEPrd" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirEPrd" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente Desativar a Especificação?')) return false;" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparEPrd" Text="Limpar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Especificação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEspecificacao" runat="server" Width="576px" Enabled="False" data-ToolTip="default" ToolTip="Código da Especificação" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Faixa inicial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFaixaInicial" runat="server" CssClass="txtDecimal2" Width="200px" Enabled="False" data-ToolTip="default" ToolTip="Faixa Inicial." />
                            </div>
                            <div class="collbl">
                                Faixa final:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFaixaFinal" runat="server" CssClass="txtDecimal2" Width="200px" Enabled="False" data-ToolTip="default" ToolTip="Faixa Final." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridEspecificacaoDoProduto" runat="server" Width="100%" ForeColor="#333333" GridLines="None"
                                CellPadding="4" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:LinkButton ID="lnkSelecionarEPrd" CssClass="lnk"
                                                data-tooltip="default" ToolTip="Selecionar registro." runat="server" Text=" &gt; "
                                                OnClick="lnkSelecionarEPrd_Click">
                                                <i class="fa fa-arrow-right seta"></i>
                                            </asp:LinkButton>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="EspecificacaoDoProduto.Descricao" HeaderText="Especificação">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="FaixaInicial" HeaderText="Faixa Inicial" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="FaixaFinal" HeaderText="Faixa Final" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Ativo" HeaderText="Ativo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel6" ID="TabPanel6">
                    <HeaderTemplate>
                        EPI 
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoEPI" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtivarEPI" Text="Ativar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirEPI" Text="Desativar" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparEPI" Text="Limpar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                EPI:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEPI" runat="server" Width="576px" Enabled="False" data-ToolTip="default" ToolTip="Código EPI" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridEPIProduto" runat="server" Width="100%" ForeColor="#333333" GridLines="None"
                                CellPadding="4" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:LinkButton ID="lnkSelecionarEPI" CssClass="lnk"
                                                data-tooltip="default" ToolTip="Selecionar EPI." runat="server" Text=" &gt; "
                                                OnClick="lnkSelecionarEPI_Click">
                                                <i class="fa fa-arrow-right seta"></i>
                                            </asp:LinkButton>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoEPI" HeaderText="CodigoEPI">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descricao">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Ativo" HeaderText="Ativo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Procedimentos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoProcedimento" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkAtivarProcedimento" Text="Ativar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirProcedimento" Text="Desativar" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparProcedimento" Text="Limpar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Procedimento:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProcedimento" runat="server" Enabled="false" AutoPostBack="True" Width="596px" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 100px;">
                            <asp:GridView ID="gridProcedimento" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:LinkButton ID="lnkSelecionarProcedimento" CssClass="lnk"
                                                data-tooltip="default" ToolTip="Selecionar Procedimento." runat="server" Text=" &gt; "
                                                OnClick="lnkSelecionarProcedimento_Click">
                                                <i class="fa fa-arrow-right seta"></i>
                                            </asp:LinkButton>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoProcedimento" HeaderText="CodigoProcedimento">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="descricao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Ativo" HeaderText="Ativo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel7" runat="server" HeaderText="TabPanel7">
                    <HeaderTemplate>
                        Centro de Custos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovoProdxCons" Text="Novo" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluirProdxCons" Text="Excluir" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparProdxCons" Text="Limpar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Centro de Custo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCentroCustoProdxCons" runat="server" Width="500px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conta:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlPlanoContasProdxCons" runat="server" Width="500px" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridProdxCons" runat="server" Width="100%" ForeColor="#333333" GridLines="None"
                                CellPadding="4" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:LinkButton ID="lnkSelecionarProdxCons" CssClass="lnk"
                                                data-tooltip="default" ToolTip="Selecionar ProdutoXConsumo." runat="server"
                                                OnClick="lnkSelecionarProdxCons_Click">
                                                <i class="fa fa-arrow-right seta"></i>
                                            </asp:LinkButton>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoProduto" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Produto.Nome" HeaderText="Produto Nome">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoCusto" HeaderText="Centro de Custo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoConta" HeaderText="Plano de Contas">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="UsuarioInclusao" HeaderText="Usuario inclusão">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="UsuarioAlteracao" HeaderText="Usuario alteração">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel4" runat="server" HeaderText="TabPanel4">
                    <HeaderTemplate>
                        Consolidar Produtos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="subtitulodiv" runat="server">
                            <asp:Label ID="lblConsolidaProduto" runat="server" />
                        </div>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkAdPrdAgrp" Text="Adicionar Produto" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPrdAgrp" runat="server" Width="400px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnSelPrdAgrp" runat="server" Text=">" CssClass="btn" />
                                <asp:HiddenField ID="HidProdAgr" runat="server" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridProdutoAgrupado" runat="server" CellPadding="4" ForeColor="#333333"
                                GridLines="None" AutoGenerateColumns="False" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:BoundField DataField="CodigoProdutoAgrupado" HeaderText="Produto" />
                                    <asp:BoundField DataField="NomeProdutoAgrupado" HeaderText="Nome" />
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
                        </div>
                        <div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel3" ID="TabPanel3" TabIndex="3">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid" style="width: 99%; margin-left: 0.5%; height: 655px;">
                            <asp:GridView ID="gridConsulta" runat="server" Width="100%" ForeColor="#333333" GridLines="None"
                                CellPadding="4" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:LinkButton ID="lnkSelecionarPrd" CssClass="lnk"
                                                data-tooltip="default" ToolTip="Selecionar registro." runat="server" Text=" &gt; "
                                                OnClick="lnkSelecionarPrd_Click">
                                                <i class="fa fa-arrow-right seta"></i>
                                            </asp:LinkButton>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="DescSituacao" HeaderText="Situacao"></asp:BoundField>
                                    <asp:BoundField DataField="CodigoGrupo" HeaderText="Grupo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nome" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NCM" HeaderText="NCM">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoProdutoTerceiro" HeaderText="CodTerceiro">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>
