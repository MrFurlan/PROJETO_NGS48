<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="MonitorDeXML.aspx.vb" Inherits="NGS.Web.UI.MonitorDeXML" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1360px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngMonitorDeXML" runat="server" EnableScriptGlobalization="True" EnableScriptLocalization="True" AsyncPostBackTimeout="5000">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlMonitorDeXML" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hidLoad" runat="server" />
            <div class="titulodiv">
                Monitor de XML
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
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
                    <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeNegocio_SelectedIndexChanged" Width="582px" />
                </div>
                <div class="painelright">
                    <div class="collbl">
                        Última atualização:
                    </div>
                    <div class="coltxt">
                        <asp:Label ID="lblAtualizado" runat="server" CssClass="primario" Style="color: blue" />
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged" Width="582px" />
                </div>
                <div class="painelright">
                    <div class="collbl">
                        Proxíma atualização:
                    </div>
                    <div class="coltxt">
                        <asp:Label ID="lblProximaAtualizacao" runat="server" CssClass="primario" Style="color: red" />
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbPendente" runat="server" Checked="True" Text="Pendente(s) de Lançamento" GroupName="TipoArquivo" data-ToolTip="default" ToolTip="Pendente(s) de Lançamento" />
                    <asp:RadioButton ID="rbLancados" runat="server" Text="Lançados" GroupName="TipoArquivo" data-ToolTip="Lançados" />
                    <asp:RadioButton ID="rbSemAnexo" runat="server" Text="Sem Anexo do XML" GroupName="TipoArquivo" data-ToolTip="Sem Anexo do XML" />
                    <asp:RadioButton ID="rbTodos" runat="server" Text="Todos" GroupName="TipoArquivo" data-ToolTip="default" ToolTip="Todos" />
                </div>
                <div class="coltxt" style="margin-left: 92px;">
                    <asp:CheckBox ID="chkNFe" runat="server" Text="NFe" data-ToolTip="default"
                        ToolTip="XML de NFe." />
                    <asp:CheckBox ID="chkCTe" runat="server" Text="CTe" data-ToolTip="default"
                        ToolTip="XML de CTe." />
                </div>
                <div class="row">
                    <div class="collbl">
                        Data:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="112px"
                            data-ToolTip="default" ToolTip="Período de pesquisa." />
                        &nbsp;<asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="120px"
                            data-ToolTip="default" ToolTip="Período de pesquisa." />
                    </div>
                </div>
                <div class="painelright" style="margin-right: 80px;">
                    <div class="coltxt">
                        <asp:Label ID="lblRegistros" runat="server" CssClass="primario" Font-Size="12pt" Font-Bold="True" />
                    </div>
                </div>
            </div>
            <div class="bordagrid" style="height: 568px;">
                <asp:GridView ID="gridXML" runat="server" Width="100%" ForeColor="#333333" GridLines="None" CellPadding="4" AutoGenerateColumns="False">
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                    <Columns>
                        <asp:BoundField DataField="Cliente" HeaderText="Fornecedor">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ClienteNome" HeaderText="Fornecedor">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ClienteCidade" HeaderText="Fornecedor">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Chave" HeaderText="Chave">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="" ShowHeader="False">
                            <ItemTemplate>
                                <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                                    <ContentTemplate>
                                        <asp:ImageButton ID="imgDownload" runat="server" ImageUrl="~/images/xml16x16.png"
                                            ImageAlign="AbsMiddle" OnCommand="imgDownload_Click" CommandName='<%# Eval("Codigo") %>' data-ToolTip="default"
                                            ToolTip="Download do arquivo" />
                                    </ContentTemplate>
                                    <Triggers>
                                        <asp:PostBackTrigger ControlID="imgDownload" />
                                    </Triggers>
                                </asp:UpdatePanel>
                            </ItemTemplate>
                            <ItemStyle Width="30px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="" ShowHeader="False">
                            <ItemTemplate>
                                <asp:ImageButton ID="XmlNFAutomatico" runat="server" ImageUrl="~/images/xmlExp16x16.png"
                                    ImageAlign="AbsMiddle" OnCommand="imgXmlNFAutomatico_Click" data-ToolTip="default"
                                    ToolTip="Enviar XML para emissão de NF" />
                            </ItemTemplate>
                            <ItemStyle Width="30px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="" ShowHeader="False">
                            <ItemTemplate>
                                <asp:ImageButton ID="PdfNFView" runat="server" ImageUrl="~/images/icopdf16X16.jpg"
                                    ImageAlign="AbsMiddle" OnCommand="imgPdfNFView_Click" data-ToolTip="default"
                                    ToolTip="Visualizar PDF NF-e" />
                            </ItemTemplate>
                            <ItemStyle Width="30px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="Tipo" HeaderText="Tipo">
                            <HeaderStyle HorizontalAlign="Left" Width="20px" />
                            <ItemStyle HorizontalAlign="Left" Width="20px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Serie" HeaderText="Série">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="nNF" HeaderText="Número">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Emissao" HeaderText="Emissao" DataFormatString="{0:dd/MM/yyyy}">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="CFOP" HeaderText="CFOP">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorDoProduto" HeaderText="Valor Produto">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorTotalDaNota" HeaderText="Valor Total">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Lancada" HeaderText="Lançada">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Usuario" HeaderText="Usuário">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="DiaLancamento" HeaderText="Data">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="TemAnexo" HeaderText="Arquivo ANEXO">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="">
                            <ItemTemplate>
                                <asp:ImageButton ID="imgInfXML" runat="server"
                                    Width="16px" Height="16px" ImageUrl="~/images/important.png" Visible="false" data-ToolTip="default"
                                    ToolTip="Cancelado." />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="">
                            <ItemTemplate>
                                <!-- Exemplo de uso de controle dentro do TemplateField -->
                                <asp:TextBox ID="visibleImageButton" runat="server" Text='<%# Eval("visibleImageButton") %>' Visible="false" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="">
                            <ItemTemplate>
                                <!-- Exemplo de uso de controle dentro do TemplateField -->
                                <asp:TextBox ID="ToolTip" runat="server" Text='<%# Eval("ToolTip") %>' Visible="false" />
                            </ItemTemplate>
                        </asp:TemplateField>
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
    </asp:UpdatePanel>
    <uc:EnviarXMLEmissao ID="ucEnviarXMLEmissao" runat="server" />
</asp:Content>
