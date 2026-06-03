<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PagamentoDeTitulos.aspx.vb" Inherits="NGS.Web.UI.PagamentoDeTitulos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadPagamentoDeTitulos() {
            if ($("#<%=ddlCarteira.ClientID%>").val() == 1 || $("#<%=ddlCarteira.ClientID%>").val() == 2 || $("#<%=ddlCarteira.ClientID%>").val() == 3 || $("#<%=ddlCarteira.ClientID%>").val() == 5) {
                $("#<%=lnkRemessa.ClientID%>").parent().show();
            }
            else {
                $("#<%=lnkRemessa.ClientID%>").parent().hide();
            }

            $("#<%=ddlCarteira.ClientID%>").change(function () {
                if ($(this).val() == 1 || $(this).val() == 2 || $(this).val() == 3 || $(this).val() == 5) {
                    $("#<%=lnkRemessa.ClientID%>").parent().show("drop");
                }
                else {
                    $("#<%=lnkRemessa.ClientID%>").parent().hide("drop");
                }
            });
        }

        $(document).ready(function () {
            pageLoadPagamentoDeTitulos();
            var prmPagamentoDeTitulos = Sys.WebForms.PageRequestManager.getInstance();
            prmPagamentoDeTitulos.add_endRequest(pageLoadPagamentoDeTitulos);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPagamentoDeTitulos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPagamentoDeTitulos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Pagamento de Títulos
            </div>
            <%--<select id="teste" on--%>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRemessa" Text="Remessa" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa Pagadora:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="670px" OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Fornecedor:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeCliente" runat="server" Width="630px" Enabled="false" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" OnClick="btnCliente_Click" CssClass="btn" runat="server"
                        Text=">" data-ToolTip="default" ToolTip="Selecionar o fornecedor desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Carteira:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCarteira" runat="server" Width="670px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pagar:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlReceber" runat="server" Width="670px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSituacao" runat="server" Width="309px" />
                </div>
                <div class="collbl">
                    Pesquisa por:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdEmissao" Checked="True" runat="server" Text=" Movimento "
                        GroupName="consulta" data-ToolTip="default" ToolTip="Informar se a consulta é por movimento ou data de vencimento." />
                    <asp:RadioButton ID="rdVenc" runat="server" Text=" Data Vencimento " GroupName="consulta"
                        data-ToolTip="default" ToolTip="Informar se a consulta é por movimento ou data de vencimento." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Inicial:
                </div>
                <div class="coltxt" style="width: 309px;">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" Width="100px" runat="server"
                        data-ToolTip="default" ToolTip="Informar o período inicial de consulta." />
                </div>
                <div class="collbl">
                    Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" Width="100px" runat="server"
                        data-ToolTip="default" ToolTip="Informar o período final de consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Total:
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblQtdeEValorTotal" runat="server" Font-Bold="True" data-ToolTip="default"
                        ToolTip="" />
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
            <div class="row" runat="server" visible="False">
                <div class="coltxt" style="margin-left: 125px;">
                    <asp:TextBox ID="txtNumValor" runat="server" Visible="False" />
                    <asp:Button ID="btnTotalTitulos" OnClick="btnTotalTitulos_Click" runat="server" Width="120px"
                        Text="Totalizar Seleção" Visible="False" CssClass="btn" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridPagamentoDeTitulos" runat="server" Width="100%" Height="185px"
                    ForeColor="#333333" EnableTheming="False" GridLines="None" CellPadding="4" AutoGenerateColumns="False"
                    AllowSorting="True" OnSelectedIndexChanged="gridPagamentoDeTitulos_SelectedIndexChanged">
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                    <Columns>
                        <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                <asp:CheckBox ID="chkTodosGridTitulos" runat="server" AutoPostBack="True" OnCheckedChanged="ChkTodosGridTitulos_CheckedChanged" />
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="ChkGridTitulos" runat="server" AutoPostBack="True" OnCheckedChanged="ChkGridTitulos_CheckedChanged"
                                    Checked='<%# eval("Enviar") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Titulo_Id" HeaderText="Registro">
                            <ItemStyle Width="60px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Venc."
                            HtmlEncode="False">
                            <ItemStyle Width="80px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="ReduzidoOrigem" HeaderText="Empresa" Visible="false" />
                        <asp:TemplateField HeaderText="Fornecedor">
                            <ItemTemplate>
                                CNPJ:&nbsp;<%# Eval("Fornecedor")%><br />
                                Nome:&nbsp;<%# Eval("NomeFornecedor")%>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Fornecedor" HeaderText="Cnpj Forn." Visible="false"></asp:BoundField>
                        <asp:BoundField DataField="NomeFornecedor" HeaderText="Nome Fornecedor" Visible="false"></asp:BoundField>
                        <asp:TemplateField HeaderText="Histórico">
                            <ItemTemplate>
                                <b>
                                    <%# Eval("DadosBancariosFornc")%></b><br />
                                Histórico:&nbsp;<%# Eval("Historico")%>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="ValorParaPagamento" HeaderText="Valor" ReadOnly="True" />
                        <asp:BoundField DataField="Historico" HeaderText="Historico" Visible="false" />
                        <asp:BoundField DataField="DadosBancariosFornc" HeaderText="Dados Banc&#225;rios/Fornec."
                            Visible="false">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Pagamento">
                            <ItemTemplate>
                                Tipo:&nbsp;<b><%# Eval("TipoPagtoDesc")%></b><br />
                                <%# Eval("CodigoDeBarras")%>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="TipoPagtoDesc" HeaderText="Desc Tipo Pgto" ReadOnly="True"
                            Visible="false"></asp:BoundField>
                        <asp:BoundField DataField="CodigoDeBarras" HeaderText="Codigo De Barras" Visible="false" />
                        <asp:BoundField DataField="DadosBancariosPagadora" HeaderText="Dados Bancarios Pagadora"
                            Visible="false">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Carteira" HeaderText="Carteira" ReadOnly="True" Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="DescCarteira" ReadOnly="True" Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="Situacao" HeaderText="C&#243;d." Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="TipoPagto" HeaderText="Tipo Pgto" ReadOnly="True" Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Mov."
                            HtmlEncode="False" Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="ReduzidoFornecedor" HeaderText="Reduzido Forn." Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="MoedaProgramado" HeaderText="Valor U$" ReadOnly="True"
                            Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="Acrescimos" HeaderText="Acrescimos" ReadOnly="True" Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="Descontos" HeaderText="Descontos" Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="Juros" HeaderText="Juros" Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="Deducoes" HeaderText="Dedu&#231;&#245;es" Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="CidadeFornecedor" HeaderText="Cidade Forn." Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="EstadoFornecedor" HeaderText="UF Forn." Visible="False"></asp:BoundField>
                        <asp:BoundField DataField="OficialProgramado" HeaderText="Valor R$" ReadOnly="True"
                            Visible="False"></asp:BoundField>
                    </Columns>
                    <EditRowStyle BackColor="#999999" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                </asp:GridView>
            </div>

        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaDadosBancarios ID="ucConsultaDadosBancarios" runat="server" />
</asp:Content>
