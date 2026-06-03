<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="ReemissaoBoletoBancario.aspx.vb" Inherits="NGS.Web.UI.ReemissaoBoletoBancario" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngReemissaoBoletoBancario" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlReemissaoBoletoBancario" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Reemissão Boleto Bancário
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkImpressao" Text="Impressão" runat="server" />
                        </li>
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
                    <asp:DropDownList ID="ddlEmpresa" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged" Width="596px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtClientes" runat="server" Width="585px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaCliente" OnClick="cmdBuscaCliente_Click" runat="server" Text=">"
                        CssClass="btn" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default"
                        ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkPeriodo" runat="server" Checked="true" Text="Período:" ToolTip="Marcar para consultar por período." />
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtDataInicial" Width="116px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Data inicial da consulta." />
                    &nbsp;&nbsp;à&nbsp;
                    <asp:TextBox runat="server" ID="txtDataFinal" Width="116px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Data final da consulta." />
                    <asp:RadioButton ID="radMovimento" runat="server" AutoPostBack="True" Text="Movimento" GroupName="radPeriodo" Checked="true" ToolTip="Consultar período por movimento." />
                    <asp:RadioButton ID="radVencimento" runat="server" AutoPostBack="True" Text="Vencimento" GroupName="radPeriodo" ToolTip="Consultar período por vencimento." />
                </div>
            </div>
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
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNotaFiscal" runat="server" data-ToolTip="default" ToolTip="Número da nota fiscal." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridBoletoBancario" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                <asp:CheckBox ID="chkAllTitulos" data-ToolTip="default" ToolTip="Seleciona todos os títulos."
                                    Text="CK" runat="server" AutoPostBack="True" OnCheckedChanged="chkAllTitulos_CheckedChanged" />
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkImpBoletos" runat="server" />
                            </ItemTemplate>
                            <ControlStyle Width="30px" />
                            <HeaderStyle HorizontalAlign="Center" Width="30px" />
                            <ItemStyle HorizontalAlign="Center" Width="30px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Registro">
                            <ItemTemplate>
                                <asp:HyperLink Style="display: none;" ID="hpTitulo" runat="server" NavigateUrl="#" Text='<%# Eval("Codigo")%>' />
                                <a href="#" onclick="newTab('<%# Eval("CodigoCifrado", "WFTitulo.aspx?param={0}")%>')"><%# Eval("Codigo")%></a>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Cidade" HeaderText="Empresa">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Histórico">
                            <HeaderStyle HorizontalAlign="Left" />
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
                                            ImageAlign="AbsMiddle" OnCommand="btnBoleto_Click" CommandName='<%# Eval("Codigo") %>' data-ToolTip="default"
                                            ToolTip="Impressão do Boleto Bancário" />
                                        <controlstyle width="1px" />
                                        <headerstyle horizontalalign="Left" width="1px" />
                                        <itemstyle horizontalalign="Left" width="1px" />
                                    </ContentTemplate>
                                    <Triggers>
                                        <asp:PostBackTrigger ControlID="imgBoletoPDF" />
                                    </Triggers>
                                </asp:UpdatePanel>
                            </ItemTemplate>
                            <ControlStyle Width="60px" />
                            <HeaderStyle HorizontalAlign="Center" Width="60px" />
                            <ItemStyle HorizontalAlign="Center" Width="60px" />
                        </asp:TemplateField>
                        <asp:TemplateField ShowHeader="False"></asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
            <asp:HiddenField ID="hidBoletoNotaFiscal" runat="server" />
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
