<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="AlterarVencimentosNF.aspx.vb" Inherits="NGS.Web.UI.AlterarVencimentosNF" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngAlterarVencimentosNF" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAlterarVencimentosNF" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Alterar Vencimentos da Nota Fiscal
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
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
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="633px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="633px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="593px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False"
                        data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    UF:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtUf" runat="server" Width="100px" MaxLength="2" data-ToolTip="default"
                        ToolTip="Abreviatura do estado." />
                </div>
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt" style="width: 110px;">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="79px"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="100px"
                        data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNota" Style="text-align: right;" runat="server" class="txtNumerico"
                        Width="100px" data-ToolTip="default" ToolTip="Número da Nota Fiscal." />
                </div>
                <div class="collbl">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtES" runat="server" Width="100px" MaxLength="1" data-ToolTip="default"
                        ToolTip="Informar se a nota é de entrada ou saída." />
                </div>
                <div class="collbl">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerie" runat="server" Width="100px" MaxLength="3" data-ToolTip="default"
                        ToolTip="Série da Nota Fiscal." />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 250px;">
                    <asp:Label runat="server" ID="txtTotal"  Text="Valor total dos titulos: " />
                    <asp:Label runat="server" ID="lblTotal" Text="" />
                </div>
                <div class="coltxt" style="width: 235px;">
                    <asp:Label runat="server" ID="txtApurado" Text="Valor total apurado: " />
                    <asp:Label runat="server" ID="lblApurado" Text="" />
                </div>
                 <div class="coltxt">
                    <asp:Label runat="server" ID="txtDiferenca" Text="Valor da diferença: " />
                    <asp:Label runat="server" ID="lblDiferenca" Text="" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridConsultaTitulos" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <Columns>
                        <asp:BoundField DataField="Codigo" HeaderText="Título">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle Width="150px" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Vencimento">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                            <ItemTemplate>
                                <asp:TextBox ID="txtVencimento" runat="server" Enabled="False" CssClass="calendario" Width="65px"
                                    Text='<%# Eval("Vencimento") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Reais">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                            <ItemTemplate>
                                <asp:TextBox ID="txtValorDoDocumento" runat="server" Enabled="False" CssClass="txtDecimal"  Width="70px"
                                    Text='<%# Eval("ValorDoDocumento", "{0:N2}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField ShowHeader="False">
                            <ItemTemplate>
                                <asp:Button ID="btnAlterarTitulo" runat="server" OnClick="btnAlterarTitulo_Click" Style="width: auto !important;"
                                    Text="Alterar" />
                                <asp:Button ID="btnSalvarTitulo" runat="server" CausesValidation="True" OnClientClick=" if(!confirm('Deseja realmente alterar o valor do titulo?')) return false;"
                                    Text="OK" Visible="False" OnClick="btnSalvarTitulo_Click" Style="width: auto !important;" />
                                <asp:Button ID="btnCancelarTitulo" runat="server" Text="Cancelar" Visible="False" Style="width: auto !important;"
                                    OnClick="btnCancelarTitulo_Click" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" Width="100px" />
                            <ItemStyle HorizontalAlign="Center" Width="100px" />
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
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
</asp:Content>

