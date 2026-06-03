<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ContratoBancario.aspx.vb" Inherits="NGS.Web.UI.ContratoBancario" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 125px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="smContratoBancario" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="upContratoBancario" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Contrato Bancario
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconConfirmar" ID="lnkConfirmar" runat="server" Text="Gerar Parcela(s)" />
                        </li>
                        <li runat="server" visible="false">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" runat="server" Text="Gravar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" runat="server" Text="Relatório" />
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
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlUnidadeDeNegocio" runat="server" Width="585px" OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged"
                        AutoPostBack="true" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" TabIndex="3" runat="server" Width="585px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Banco:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlBanco" runat="server" Width="585px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Contrato:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContratoBanco" runat="server" Width="100px" AutoPostBack="true"
                        data-ToolTip="default" ToolTip="Informar o numero do contrato." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Finalidade Financeira:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlFinalidade" runat="server" Width="585px" OnSelectedIndexChanged="ddlFinalidade_SelectedIndexChanged"
                        AutoPostBack="true" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Encargo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEncargo" runat="server" Width="585px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Qtde Parcelas:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtParcelas" runat="server" Width="86px" CssClass="txtNumerico"
                        data-ToolTip="default" ToolTip="Informar o número das parcelas de pagamento." />
                </div>
                <div class="collbl">
                    Valor:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValor" runat="server" Width="100px" CssClass="txtDecimal" Enabled="false"
                        data-ToolTip="default" ToolTip="Valor total do título." />
                </div>
                <div class="collbl">
                    Movimento:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtMovimento" runat="server" Width="86px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Data do lançamento." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridParcBancario" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                            <ItemStyle Width="25px"></ItemStyle>
                        </asp:CommandField>
                        <asp:TemplateField HeaderText="Vencimento" HeaderStyle-Width="120px">
                            <ItemTemplate>
                                <asp:TextBox ID="txtVencimento" runat="server" CssClass="calendario" Width="86px"
                                    Text='<%# Eval("Vencimento", "{0:dd/MM/yyyy}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="ValorDoDocumento" DataFormatString="{0:N}" HeaderText="Valor da Parcela"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Center" Width="250px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="NrParcela" HeaderText="Nr Parcela">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
