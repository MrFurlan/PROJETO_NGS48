<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="MemorandoDeExportacao.aspx.vb" Inherits="NGS.Web.UI.MemorandoDeExportacao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1024px !important;
        }

        .collbl {
            width: 135px;
        }

        .w113 {
            width: 113px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Memorando de Exportação
            </div>
            <ajaxToolkit:TabContainer ID="TBMemorando" runat="server" ActiveTabIndex="1"
                Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton ID="lnkConsulta" CssClass="iconConsultar" Text="Consultar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton ID="lnkLimpar" CssClass="iconAtualizar" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton ID="lnkAjuda" CssClass="iconAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                <asp:CheckBox ID="chkConsolidarEmitente" runat="server" Text="Cons. Emitente:" ToolTip="Consolidar Emitente" />
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="HEmitenteConsulta" runat="server" />
                                <asp:TextBox ID="txtEmitenteConsulta" runat="server" Width="480px" Enabled="false" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnEmitente" OnClick="btnEmitente_Click" runat="server" Text=">" CssClass="btn" ToolTip="Consulta Emitente." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                <asp:CheckBox ID="chkConsolidaComprovando" runat="server" Font-Bold="True" Text="Cons. Comprovando:" ToolTip="Consolidar Emitente." />
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="HComprovandoConsulta" runat="server" />
                                <asp:TextBox ID="txtComprovandoConsulta" runat="server" Width="480px" Enabled="false" />
                            </div>

                            <div class="coltxt">
                                <asp:Button ID="btnComprovandoConsulta" OnClick="btnComprovandoConsulta_Click" runat="server" CssClass="btn"
                                    Text=">" ToolTip="Consulta Comprovando." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtProdutoConsulta" runat="server" Width="480px" Enabled="false" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnConsultaProduto" OnClick="btnConsultaProduto_Click" runat="server"
                                    Text=">" CssClass="btn" ToolTip="Consulta Produto." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data Emissão:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="86px" />
                            </div>
                            <div class="coltxt">
                                á
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="86px" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridConsulta" runat="server" OnSelectedIndexChanged="gridConsulta_SelectedIndexChanged"
                                AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None"
                                Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                                    <asp:TemplateField HeaderText="Imprimir">
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgImpressao" runat="server" Height="20px" ImageUrl="~/images/impressora.JPG"
                                                OnClick="imgImpressao_Click" />
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Emitente" HeaderText="Emitente" />
                                    <asp:BoundField DataField="EndEmitente" />
                                    <asp:BoundField DataField="NomeEmitente" />
                                    <asp:BoundField DataField="CidadeUFEmitente" HeaderText="Cidade-UF" />
                                    <asp:BoundField DataField="Memorando" HeaderText="Memorando">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataMemorando" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data" />
                                    <asp:BoundField DataField="Quantidade" DataFormatString="{0:N0}" HeaderText="Qtde">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Cadastro
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton ID="lnkNovo" class="iconNovo" runat="server" Text="Gravar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton ID="lnkAlterar" class="iconAtualizar" runat="server" Text="Alterar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton ID="lnkExcluir" class="iconExcluir" runat="server" Text="Excluir" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton ID="lnkLimparCadastro" class="iconLimpar" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton ID="lnkRelatorio" class="iconRelatorio" runat="server" Text="Relatório" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton ID="lnkAjuda1" class="iconAjuda" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 57%;">
                            <div class="row">
                                <div class="collbl">
                                    Emitente:
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="HEmitente" runat="server" />
                                    <asp:TextBox ID="txtEmitente" runat="server" Width="375px" Enabled="false" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnCadEmitente" runat="server" OnClick="btnCadEmitente_Click" Text="&gt;" CssClass="btn"
                                        UseSubmitBehavior="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Exportador Equiparado:
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="HExportadorEquiparado" runat="server" />
                                    <asp:TextBox ID="TxtExpEquiparado" runat="server" Width="375px" Enabled="false" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="BtnCadExpEquiparado" OnClick="btnExpEquiparado_Click" runat="server"
                                        UseSubmitBehavior="False" Text=">" CssClass="btn" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Comprovando:                           
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="HComprovando" runat="server" />
                                    <asp:TextBox ID="txtComprovando" runat="server" Width="375px" Enabled="false" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnComprovando" OnClick="btnComprovando_Click" runat="server" UseSubmitBehavior="False" CssClass="btn"
                                        Text=">" />
                                </div>
                            </div>

                            <div class="row">
                                <div class="collbl">
                                    Mem./Equip. Export:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtMemorando" runat="server" OnTextChanged="txtMemorando_TextChanged"
                                        AutoPostBack="True" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtMemorandoEquiparadoExp" runat="server" Enabled="False" OnTextChanged="txtMemorando_TextChanged"
                                        AutoPostBack="True" Visible="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Data Memorando:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataMemorando" CssClass="calendario" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Tipo Documento:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlTipoDocumento" runat="server">
                                        <asp:ListItem Value="0">Declara&#231;&#227;o De Exporta&#231;&#227;o</asp:ListItem>
                                        <asp:ListItem Value="1">Declara&#231;&#227;o Simplificada De Exporta&#231;&#227;o</asp:ListItem>
                                    </asp:DropDownList>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Nº Despacho:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNumeroDespacho" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Data Despacho:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataDespacho" CssClass="calendario" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Nº Conh./Data/Tipo:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNumeroConhecimento" Width="70px" runat="server" />
                                </div>
                                <div class="coltxt">
                                    /
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataConhecimento" ClientIDMode="Static" runat="server" Width="70px"
                                        AutoPostBack="True" CssClass="calendario" />
                                </div>
                                <div class="coltxt">
                                    /
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlTipoConhec" runat="server" Width="118px">
                                        <asp:ListItem Value="01">AWB</asp:ListItem>
                                        <asp:ListItem Value="02">MAWB</asp:ListItem>
                                        <asp:ListItem Value="03">HAWB</asp:ListItem>
                                        <asp:ListItem Value="04">COMAT</asp:ListItem>
                                        <asp:ListItem Value="06">R. EXPRESSAS</asp:ListItem>
                                        <asp:ListItem Value="07">ETIQ. REXPRESSAS</asp:ListItem>
                                        <asp:ListItem Value="08">HR. EXPRESSAS</asp:ListItem>
                                        <asp:ListItem Value="09">AV7</asp:ListItem>
                                        <asp:ListItem Value="10">BL</asp:ListItem>
                                        <asp:ListItem Value="11">MBL</asp:ListItem>
                                        <asp:ListItem Value="12">HBL</asp:ListItem>
                                        <asp:ListItem Value="13">CRT</asp:ListItem>
                                        <asp:ListItem Value="14">DSIC</asp:ListItem>
                                        <asp:ListItem Value="16">COMAT BL</asp:ListItem>
                                        <asp:ListItem Value="17">RWB</asp:ListItem>
                                        <asp:ListItem Value="18">HRWB</asp:ListItem>
                                        <asp:ListItem Value="19">TIF/DTA</asp:ListItem>
                                        <asp:ListItem Value="20">CP2</asp:ListItem>
                                        <asp:ListItem Value="91">N&#194;O IATA</asp:ListItem>
                                        <asp:ListItem Value="92">MNAO IATA</asp:ListItem>
                                        <asp:ListItem Value="93">HNAO IATA</asp:ListItem>
                                        <asp:ListItem Value="99">OUTROS</asp:ListItem>
                                    </asp:DropDownList>
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnAddConhecimento" runat="server" UseSubmitBehavior="False" Text="Adicionar" CssClass="botao" />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="ImgAddConhec" OnClick="ImgAddConhec_Click" runat="server" Visible="False"
                                        Width="20px" Height="20px" ImageUrl="~/images/detalhes.png" ImageAlign="AbsMiddle" />
                                </div>
                            </div>
                            <div class="bordagrid" style="height: 155px; width: 99%;">
                                <asp:GridView ID="GridConhecimento" runat="server" OnPageIndexChanging="gridNotas_PageIndexChanging"
                                    AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None"
                                    OnSelectedIndexChanged="gridConsulta_SelectedIndexChanged" Width="100%">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <Columns>
                                        <asp:BoundField DataField="NumConhecimentoDeEmbarque" HeaderText="Num. Conhecimento"></asp:BoundField>
                                        <asp:BoundField DataField="DataConhecimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Conhecimento"></asp:BoundField>
                                        <asp:BoundField DataField="TipoConhecimento" HeaderText="Cod Tipo Conhecimento" Visible="False"></asp:BoundField>
                                        <asp:BoundField DataField="DescTipoConhecimento" HeaderText="Tipo Conhecimento"></asp:BoundField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:ImageButton ID="ImgExcluirConhec" runat="server" Height="20px" Visible="false"
                                                    ImageAlign="Middle" ImageUrl="~/images/erro.jpg" Width="20px" OnClick="ImgExcluirConhec_Click" />
                                                <asp:Button ID="btnExcluirConhecimento" runat="server" Text="Excluir" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 42.5%;">
                            <div class="subtitulodiv">
                                Nota Fiscal
                            </div>
                            <div class="row">
                                <div class="collbl w113">
                                    Nota(s):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtProcura" runat="server" data-ToolTip="default" ToolTip="Se for varias notas use a virgula para separa-las"
                                        Width="200px" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnConsultaNota" runat="server" UseSubmitBehavior="False" Text="consultar" CssClass="btn" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w113">
                                    Nota(s):
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtProcuraMaisNotas" runat="server" data-ToolTip="default" ToolTip="Se for varias notas use a virgula para separa-las"
                                        Width="200px" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnBuscaMaisNotas" runat="server" UseSubmitBehavior="False" Text="Adicionar" CssClass="btn" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w113">
                                    Safra:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlSafraSelecaoNotas" runat="server"
                                        Width="288px">
                                    </asp:DropDownList>
                                </div>
                            </div>
                            <div class="bordagrid" style="height: 317px;">
                                <asp:GridView ID="gridNotaSaida" runat="server" OnPageIndexChanging="gridNotaSaida_PageIndexChanging"
                                    ForeColor="#333333" PageSize="16" GridLines="None" CellPadding="4" AutoGenerateColumns="False"
                                    AllowPaging="True" OnSelectedIndexChanged="gridNotaSaida_SelectedIndexChanged">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <Columns>
                                        <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                                        <asp:BoundField DataField="NumeroNota" HeaderText="Nota"></asp:BoundField>
                                        <asp:BoundField DataField="Serie" HeaderText="Serie"></asp:BoundField>
                                        <asp:BoundField DataField="QtdeNota" DataFormatString="{0:N0}" HeaderText="Qtde Nota">
                                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundField>
                                        <asp:BoundField DataField="QtdeJaComprovada" DataFormatString="{0:N0}" HeaderText="Qtde Comprovada">
                                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                                <asp:GridView ID="gridNotas" runat="server" OnPageIndexChanging="gridNotas_PageIndexChanging"
                                    AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None"
                                    OnSelectedIndexChanged="gridConsulta_SelectedIndexChanged" Width="100%">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <Columns>
                                        <asp:BoundField DataField="NumeroNota" HeaderText="Nota"></asp:BoundField>
                                        <asp:BoundField DataField="Serie" HeaderText="Serie"></asp:BoundField>
                                        <asp:BoundField DataField="Saldo" DataFormatString="{0:N0}" HeaderText="Saldo">
                                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:Button ID="btnQtde" OnClick="btnQtde_Click" runat="server" Text=">" UseSubmitBehavior="False" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Qtde">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtQtde" runat="server" Width="92px" Text='<%# eval("QuantidadeMemorando") %>'
                                                    OnTextChanged="txtQtde_TextChanged" AutoPostBack="True" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:Button ID="btnZerar" runat="server" OnClick="btnZerar_Click" Text=" X " />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Dta Averbação Declar. Export.:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataAverba" CssClass="calendario" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                País Destino:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlPais" runat="server" Width="375px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Navio:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNavio" runat="server" Width="187px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Representante:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="HRepresentante" runat="server" />
                                <asp:TextBox ID="txtRepresentante" runat="server" Width="375px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnRepresentante" OnClick="btnRepresentante_Click" runat="server" CssClass="btn"
                                    UseSubmitBehavior="False" Text=">" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtProduto" runat="server" Width="281px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnProduto" OnClick="btnProduto_Click" runat="server" UseSubmitBehavior="False" CssClass="btn"
                                    Text=">" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Nota/Serie/QtdeSaldoNota:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNota" runat="server" Enabled="False" Width="80px" />
                            </div>
                            <div class="coltxt">
                                /
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSerie" runat="server" Width="80px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                /
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQuantidadeSaldoNota" Width="80px" runat="server" Enabled="False"
                                    BackColor="#FFFF80" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="BtnNotaDeComprovacao" OnClick="BtnNotaDeComprovacao_Click" runat="server" CssClass="btn"
                                    Text=">" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Moeda/Indexador:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlMoeda" runat="server" Width="160px" />
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlIndexador" runat="server" Width="160px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Valor Nota/Data/QtdeMem.:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtValorNota" Width="100px" onkeyup="formatavalor(this, 18, 2);"
                                    runat="server" />
                            </div>
                            <div class="coltxt">
                                /
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtDataDaNota" Width="80px" runat="server" CssClass="calendario" />
                            </div>
                            <div class="coltxt">
                                /
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQuantidadeMemorando" Width="80px" runat="server" BackColor="#FFFF80"
                                    Enabled="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                R.E./Data/UFProd.   Fabric.:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtRegExportacao" runat="server" Width="80px" />
                            </div>
                            <div class="coltxt">
                                /
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtDataRegExp" CssClass="calendario" runat="server" Width="77px" />
                            </div>
                            <div class="coltxt">
                                /
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEstado" runat="server" Width="120px"
                                    AutoPostBack="True" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnAddRegExportacao" runat="server" UseSubmitBehavior="False" Text="Adicionar" CssClass="botao" />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="ImgAddRegExp" OnClick="ImgAddRegExp_Click1" runat="server" Visible="False"
                                    Width="20px" Height="20px" ImageUrl="~/images/detalhes.png" ImageAlign="AbsMiddle" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 150px;">
                            <asp:GridView ID="GridRegistroExp" runat="server" OnPageIndexChanging="gridNotas_PageIndexChanging"
                                AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None"
                                OnSelectedIndexChanged="gridConsulta_SelectedIndexChanged" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <Columns>
                                    <asp:BoundField DataField="CodRegistroDeExportacao" HeaderText="Registro De Exporta&#231;&#227;o"></asp:BoundField>
                                    <asp:BoundField DataField="DataRegExportacao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Reg. Exportacao"></asp:BoundField>
                                    <asp:BoundField DataField="UfProdutor" HeaderText="UF Produtor / Fabricante"></asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="ImgExcluirRegExp" Visible="false" runat="server" Height="20px"
                                                ImageAlign="Middle" ImageUrl="~/images/erro.jpg" Width="20px" OnClick="ImgExcluirRegExp_Click" />
                                            <asp:Button ID="btnExcluirRegExp" runat="server" Text="Excluir" OnClick="btnExcluirRegExp_Click" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <div class="subtitulodiv">
                            Drawbak
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Num.AtoConcessório:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtNumAtoConcessorio" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data Registro/Validade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtDtaRegAtoConcessorio" runat="server" CssClass="calendario" Width="77px" />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtDtaValidAtoConcessorio" runat="server" CssClass="calendario"
                                    Width="77px" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel3" runat="server">
                    <HeaderTemplate>
                        Pendências
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="collbl">
                                Filtros:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbTodosMem" runat="server" Text="Todos" Checked="True" GroupName="memo" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbEmissaoMem" runat="server" Text="Comprovação a Emitir"
                                    GroupName="memo" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbComprovacaoMem" runat="server" Text="Comprovação a Receber"
                                    GroupName="memo" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                <asp:Label ID="lblSafra" runat="server" Text="Safra :" Font-Bold="True" />
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSafra" runat="server" Width="418px">
                                </asp:DropDownList>
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnAtualizarConsulta" OnClick="btnAtualizarConsulta_Click" runat="server" CssClass="botao"
                                    Text="Cons. Pendências" />
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            Memorandos pendentes
                        </div>
                        <div class="bordagrid" style="height: 155px;">
                            <asp:GridView ID="gridClientesPendentes" runat="server" OnSelectedIndexChanged="gridClientesPendentes_SelectedIndexChanged"
                                AutoGenerateColumns="False" PageSize="16" CellPadding="4" ForeColor="#333333"
                                GridLines="None" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <Columns>
                                    <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                                    <asp:TemplateField HeaderText="Imprimir">
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgImpressaoPend" runat="server" Height="20px" ImageUrl="~/images/impressora.JPG"
                                                OnClick="imgImpressaoPend_Click" />
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Incluir">
                                        <ItemTemplate>
                                            <asp:Button ID="btnIncluirDoGrid" runat="server" Text=" + " OnClick="btnIncluirDoGrid_Click" />
                                        </ItemTemplate>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="DescSituacao" HeaderText="Situacao"></asp:BoundField>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente"></asp:BoundField>
                                    <asp:BoundField DataField="EndCliente"></asp:BoundField>
                                    <asp:BoundField DataField="NomeCliente">
                                        <HeaderStyle Width="250px"></HeaderStyle>
                                        <ItemStyle Width="250px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ES" HeaderText="E/S"></asp:BoundField>
                                    <asp:BoundField DataField="Cidade" HeaderText="Cidade"></asp:BoundField>
                                    <asp:BoundField DataField="Produto" HeaderText="Produto"></asp:BoundField>
                                    <asp:BoundField DataField="NomeProduto"></asp:BoundField>
                                    <asp:BoundField DataField="Saldo" DataFormatString="{0:N0}" HeaderText="Saldo">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <div class="subtitulodiv">
                            Notas pendentes do cliente
                        </div>
                        <div class="bordagrid" style="height: 155px;">
                            <asp:GridView ID="gridNotasPendentes" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridConsulta_SelectedIndexChanged"
                                Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <Columns>
                                    <asp:BoundField DataField="Nota" HeaderText="Nota"></asp:BoundField>
                                    <asp:BoundField DataField="Serie" HeaderText="Serie"></asp:BoundField>
                                    <asp:BoundField DataField="DataDaNota" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data"></asp:BoundField>
                                    <asp:BoundField DataField="DiasDecorridos" HeaderText="Dias Emissao "></asp:BoundField>
                                    <asp:BoundField DataField="QuantidadeNota" DataFormatString="{0:N0}" HeaderText="Qtde Nota">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="QtdeDevolvida" DataFormatString="{0:N0}" HeaderText="Qtde Devolvida">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="QtdeJaComprovada" DataFormatString="{0:N0}" HeaderText="Qtde Comprovada">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Saldo" DataFormatString="{0:N0}" HeaderText="Saldo">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <div class="subtitulodiv">
                            Notas de exportação com saldo para comparação
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:ImageButton ID="imgImpPenComprov" OnClick="imgImpPenComprov_Click" runat="server"
                                    Height="20px" ImageUrl="~/images/impressora.JPG" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 155px;">
                            <asp:GridView ID="gridNotasComprovacaoPendentes" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridConsulta_SelectedIndexChanged"
                                Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>
