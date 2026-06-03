<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="EmissaoDeDuplicatas.aspx.vb" Inherits="NGS.Web.UI.EmissaoDeDuplicatas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngDuplicataRevenda" runat="server" AsyncPostBackTimeout="1000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlDuplicataRevenda" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Emissão de Duplicatas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
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
                    Formato:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdGrafica" runat="server" Checked="True" GroupName="relatorio"
                        Text="Gráfica" data-ToolTip="default" ToolTip="Seleiconar o formato do relatório." />
                    <asp:RadioButton ID="rdMatricial" runat="server" GroupName="relatorio" Text="Matricial"
                        data-ToolTip="default" ToolTip="Seleiconar o formato do relatório." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Enabled="False" Width="585px" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=">" OnClick="btnCliente_Click" UseSubmitBehavior="False"
                        CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Enabled="False" />
                    <asp:Button ID="btnPedido" CssClass="btn" runat="server" Text=">" OnClick="btnPedido_Click"
                        UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Número do pedido." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridCondicoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField ShowHeader="False">
                            <ItemTemplate>
                                <asp:Button ID="Button1" runat="server" CausesValidation="False" CommandName="Select"
                                    Enabled='<%# Eval("Provisao") = 2 %>' Text="&gt;&gt;" UseSubmitBehavior="False" />
                            </ItemTemplate>
                            <ControlStyle Width="30px" />
                            <HeaderStyle HorizontalAlign="Center" Width="30px" />
                            <ItemStyle HorizontalAlign="Center" Width="30px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="TipoFinanceiro" HeaderText="Pagar/Receber">
                            <HeaderStyle HorizontalAlign="Center" Width="100px" />
                            <ItemStyle HorizontalAlign="Center" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Codigo" HeaderText="Titulo">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Provisao" HeaderText="Provisao">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="100px" />
                            <ItemStyle HorizontalAlign="Left" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="DataBaixa" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Pagamento"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="DescMoeda" HeaderText="Moeda"></asp:BoundField>
                        <asp:BoundField DataField="DocumentoOficial" DataFormatString="{0:n2}" HeaderText="Doc. Oficial"></asp:BoundField>
                        <asp:BoundField DataField="AcrescimoOficial" DataFormatString="{0:N2}" HeaderText="Acres. Oficial">
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="DeducaoOficial" DataFormatString="{0:N2}" HeaderText="Dedução Oficial">
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="LiquidoOficial" DataFormatString="{0:N2}" HeaderText="Liq. Oficial">
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="DocumentoMoeda" DataFormatString="{0:N2}" HeaderText="Doc. Moeda">
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="AcrescimoMoeda" DataFormatString="{0:N2}" HeaderText="Acres. Moeda">
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="DeducaoMoeda" DataFormatString="{0:N2}" HeaderText="Deduções Moeda">
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:BoundField>
                        <asp:BoundField DataField="LiquidoMoeda" DataFormatString="{0:N2}" HeaderText="Liq. Moeda">
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
