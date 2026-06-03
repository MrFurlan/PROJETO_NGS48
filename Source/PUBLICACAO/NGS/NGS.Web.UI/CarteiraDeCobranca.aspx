<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CarteiraDeCobranca.aspx.vb" Inherits="NGS.Web.UI.CarteiraDeCobranca" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngCarteiraDeCobranca" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCarteiraDeCobranca" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Carteira de Cobrança
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consulta" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkRemessa" Text="Remessa" runat="server" />
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
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Width="627px" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeCliente" runat="server" Width="587px" data-ToolTip="default"
                                    ToolTip="Selecionar o cliente desejado." />
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnCliente" runat="server" Text=">" OnClick="btnCliente_Click" CssClass="btn"
                                    data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Carteira:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCarteira" runat="server" Width="627px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Receber:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlReceber" runat="server" Width="627px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Situação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSituacao" runat="server" Width="309px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" data-ToolTip="default"
                                    ToolTip="Período inicial da consulta." />
                                &nbsp;
                                <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" data-ToolTip="default"
                                    ToolTip="Período final da consulta." />
                                &nbsp;
                                <asp:RadioButton ID="rdEmissao" runat="server" GroupName="consulta" Text="Data Emissão" />
                                &nbsp;
                                <asp:RadioButton ID="rdVenc" runat="server" Checked="True" GroupName="consulta" Text="Data Vencimento" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Relatorio:
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="chkConsolidarCliente" runat="server" Text="Consolidar Cliente"
                                    data-ToolTip="default" ToolTip="Marcar se for consolidar cliente ou não agrupar carteiras." />
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="CkRelSimples" runat="server" Text="Não Agrupar Carteiras No Relatorio"
                                    data-ToolTip="default" ToolTip="Marcar se for consolidar cliente ou não agrupar carteiras." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Ordem:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbOrdRegistro" runat="server" Checked="True" GroupName="Ordem"
                                    Text="Registro" data-ToolTip="default" ToolTip="Selecionar a opção de ordem do relatório." />
                                &nbsp;
                                <asp:RadioButton ID="rbOrdSequencia" runat="server" GroupName="Ordem" Text="Sequencia"
                                    data-ToolTip="default" ToolTip="Selecionar a opção de ordem do relatório." />
                                &nbsp;
                                <asp:RadioButton ID="rbOrdVencimento" runat="server" GroupName="Ordem" Text="Vencimento"
                                    data-ToolTip="default" ToolTip="Selecionar a opção de ordem do relatório." />
                                &nbsp;
                                <asp:RadioButton ID="rbOrdCliente" runat="server" GroupName="Ordem" Text="Cliente"
                                    data-ToolTip="default" ToolTip="Selecionar a opção de ordem do relatório." />
                                &nbsp;
                                <asp:RadioButton ID="rbOrdValor" runat="server" GroupName="Ordem" Text="Valor" data-ToolTip="default"
                                    ToolTip="Selecionar a opção de ordem do relatório." />
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
                        <div class="bordagrid" style="height: auto;">
                            <asp:HiddenField ID="HfCarteiraDeCobrancaIndex" runat="server" />
                            <asp:GridView ID="gridCarteiraDeCobranca" runat="server" AllowSorting="True" AutoGenerateColumns="False"
                                CellPadding="4" EnableTheming="False" ForeColor="#333333" GridLines="None" Height="1px"
                                OnSelectedIndexChanged="gridCarteiraDeCobranca_SelectedIndexChanged" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkGridTitulos" runat="server" AutoPostBack="True" Checked='<%#Eval("Enviar") %>'
                                                OnCheckedChanged="chkGridTitulos_CheckedChanged" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Nota_Id" HeaderText="Nota" ReadOnly="True" >
                                        <ItemStyle Width="160px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Registro_Id" HeaderText="Registro">
                                        <ItemStyle Width="60px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Venc."
                                        HtmlEncode="False">
                                        <ItemStyle Width="80px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Prorrogacao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Prorrogacao" ReadOnly="True" />
                                    <asp:BoundField DataField="Sacado" HeaderText="Sacado">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Receber" HeaderText="Valor" />
                                    <asp:BoundField DataField="Carteira_Id" HeaderText="Carteira_Id" ReadOnly="True"
                                        Visible="False" />
                                    <asp:BoundField DataField="DescSituacao" HeaderText="Situacao" />
                                    <asp:BoundField DataField="NomeSacado" HeaderText="Cliente" ReadOnly="True" />
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <EmptyDataTemplate>
                                    
                                </EmptyDataTemplate>
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                    <HeaderTemplate>
                        Manutenção / Remessa
                    </HeaderTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
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
                                        </ul>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Arquivo:
                                    </div>
                                    <div class="coltxt">
                                        <asp:FileUpload ID="fup" ClientIDMode="Static" runat="server" Width="575px" onchange="this.form.submit();" />
                                        <asp:Button ID="btnProcessarRetorno" runat="server" Style="display: none" />
                                        <div id="divRetorno" runat="server" Visible="false">
                                            <asp:Label ID="lblArquivoRetorno" runat="server"></asp:Label>
                                            <asp:ImageButton ID="imbExcluirRetorno" OnClick="imbExcluirRetorno_Click" ToolTip="Remover arquivo" ImageUrl="~/Images/borracha.jpg"  runat="server" />
                                        </div>
                                    </div>
                                </div>
                                <div class="bordagrid">
                                    <asp:GridView ID="GridRetornoTitulos" runat="server" AutoGenerateColumns="False"
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
                                                    <asp:HyperLink ID="hpTitulo" runat="server" Target="_blank" NavigateUrl='<%# Eval("CodigoCifrado", "WFTitulo.aspx?param={0}")%>' Text='<%# Eval("Codigo")%>' />
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="Historico" HeaderText="Histórico">
                                                <HeaderStyle HorizontalAlign="Left" Width="380px" />
                                                <ItemStyle HorizontalAlign="Left" Width="380px"/>
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Observacoes" HeaderText="Situação">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ValorLiquido" DataFormatString="{0:n2}" HeaderText="Valor">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </ContentTemplate>
                            <%--<Triggers>
                                <asp:PostBackTrigger ControlID="lnkImportar" />
                            </Triggers>--%>
                        </asp:UpdatePanel>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
