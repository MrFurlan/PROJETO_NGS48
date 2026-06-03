<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="DuplicataAvulsa.aspx.vb" Inherits="NGS.Web.UI.DuplicataAvulsa" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngDuplicataAvulsa" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlDuplicataAvulsa" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Duplicata Avulsa
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkEmitir" Text="Emitir" runat="server" />
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
                        Width="601px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="601px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtClientes" runat="server" Width="562px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="ddlBuscaCliente" OnClick="ddlBuscaCliente_Click" runat="server" Text=">"
                        CssClass="btn" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default"
                        ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Praça Pgto:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoPraca" runat="server" />
                    <asp:TextBox ID="txtPraca" runat="server" Width="562px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultaPraca" runat="server" OnClick="btnConsultaPraca_Click"
                        CssClass="btn" Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Informações para depósito bancário." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumNota" runat="server" Width="107px" data-ToolTip="default"
                        ToolTip="Número da Nota Fiscal." />
                </div>
                <div class="collbl">
                    Duplicata:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDuplicata" runat="server" Width="107px" data-ToolTip="default"
                        ToolTip="Numero da Duplicata." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Parcelas:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtParcelas" Width="107px" runat="server" OnTextChanged="txtTotalNota_TextChanged"
                        data-ToolTip="default" ToolTip="Número de parcelas." />
                </div>
                <div class="collbl">
                    Data base Venct:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataVenc" CssClass="calendario" runat="server" Width="86px" OnTextChanged="txtTotalNota_TextChanged"
                        data-ToolTip="default" ToolTip="Data do vencimento da duplicata." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Total da Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTotalNota" CssClass="txtDecimal" runat="server" AutoPostBack="True"
                        Width="107px" OnTextChanged="txtTotalNota_TextChanged" data-ToolTip="default"
                        ToolTip="Valor total da Nota Fiscal." />
                </div>
                <div class="collbl">
                    Data Emissão:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataEmissao" CssClass="calendario" runat="server" Width="86px"
                        data-ToolTip="default" ToolTip="Data de criação da duplicata." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Vencimento:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" runat="server" CssClass="calendario" Width="86px" data-ToolTip="default"
                        ToolTip="Data de vencimento para pagamento." />
                </div>
                <div class="collbl">
                    Valor Parc:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValor" CssClass="txtDecimal" Width="107px" runat="server" AutoPostBack="True"
                        TabIndex="14" data-ToolTip="default" ToolTip="Valor de cada parcela." />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btAlterar" runat="server" OnClick="Alterar_Click" Text="Alterar"
                        CssClass="botao" UseSubmitBehavior="False" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridParcelas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridParcelas_SelectedIndexChanged">
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
                                    Text=">>" />
                            </ItemTemplate>
                            <ControlStyle Width="30px" />
                            <HeaderStyle HorizontalAlign="Center" Width="30px" />
                            <ItemStyle HorizontalAlign="Center" Width="30px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="100px" />
                            <ItemStyle HorizontalAlign="Left" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Valor" DataFormatString="{0:n2}" HeaderText="Valor Parcela"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" Width="150px" />
                            <ItemStyle HorizontalAlign="Right" Width="150px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Saldo" DataFormatString="{0:n2}" HeaderText="Saldo" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" Width="150px" />
                            <ItemStyle HorizontalAlign="Right" Width="150px" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
