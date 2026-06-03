<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="FolhaPagamento.aspx.vb" Inherits="NGS.Web.UI.FolhaPagamento" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngBoletoBancarioNotaFiscal" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:ajaxupdating id="ajaxUpdating" runat="server" text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlBoletoBancarioNotaFiscal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <ajaxtoolkit:tabcontainer id="TabContainer1" runat="server" activetabindex="0" width="100%">
                <ajaxtoolkit:tabpanel id="TabPanel3" runat="server" headertext="TabPanel3">
                    <headertemplate>
                        Vencimentos
                    </headertemplate>
                    <contenttemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkGerarRemessa" runat="server" Text="Gerar Arquivo de Remessa" />
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
                                            <asp:Button ID="btnDownload" runat="server" data-ToolTip="default" ToolTip="Baixar arquivo de folha de pagamentos"
                                                Text="Baixar arquivo de folha de pagamentos" Visible="false" />
                                        </ContentTemplate>
                                        <Triggers>
                                            <asp:PostBackTrigger ControlID="btnDownload" />
                                        </Triggers>
                                    </asp:UpdatePanel>
                                </div>
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridFuncionarios" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField HeaderText="Registro">
                                        <ItemTemplate>
                                            <asp:HyperLink Style="display: none;" ID="hpTitulo" runat="server" NavigateUrl="#" Text='<%# Eval("Codigo")%>' />
                                            <a href="#" onclick="newTab('<%# Eval("CodigoCifrado", "WFTitulo.aspx?param={0}")%>')"><%# Eval("Codigo")%></a>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="UnidadeDeNegocio" HeaderText="Unidade">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Movimento" HeaderText="Movimento">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Carteira" HeaderText="Carteira">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <%-- <asp:BoundField DataField="Encargo" HeaderText="Encargo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>--%>
                                    <asp:BoundField DataField="Cliente" HeaderText="CPF">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nome" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Historico" HeaderText="Histórico">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorLiquido" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </contenttemplate>
                </ajaxtoolkit:tabpanel>
                <ajaxtoolkit:tabpanel id="TabPanel4" runat="server">
                    <headertemplate>
                        Processa Retorno Folha de Pagamento
                    </headertemplate>
                    <contenttemplate>
                        <asp:UpdatePanel ID="upRetornoBancario" runat="server">
                            <contenttemplate>
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
                                    <asp:TemplateField HeaderText="Registro">
                                        <ItemTemplate>
                                            <asp:HyperLink Style="display: none;" ID="hpTitulo" runat="server" NavigateUrl="#" Text='<%# Eval("Codigo")%>' />
                                            <a href="#" onclick="newTab('<%# Eval("CodigoCifrado", "WFTitulo.aspx?param={0}")%>')"><%# Eval("Codigo")%></a>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="UnidadeDeNegocio" HeaderText="Unidade">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Movimento" HeaderText="Movimento">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente" HeaderText="CPF">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nome" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Prorrogacao" HeaderText="Data Pagto">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorLiquido" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                                    </asp:GridView>
                                </div>
                            </contenttemplate>
                        </asp:UpdatePanel>
                    </contenttemplate>
                </ajaxtoolkit:tabpanel>
            </ajaxtoolkit:tabcontainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:consultaclientes id="ucConsultaClientes" runat="server" />
</asp:Content>
