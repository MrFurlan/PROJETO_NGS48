<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="DesvincularRomaneioNotaFiscal.aspx.vb" Inherits="NGS.Web.UI.DesvincularRomaneioNotaFiscal" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmDesvincularRomaneioNotaFiscal" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlDesvincularRomaneioNotaFiscal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Desvincular Romaneio da Nota Fiscal
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li id="Li1" runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server" Text="Atualizar" />
                        </li>
                        <li id="Li2" runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li id="Li3" runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li id="Li4" runat="server">
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
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtES" runat="server" Width="100px" MaxLength="1" data-ToolTip="default"
                        ToolTip="Informar se a nota é de entrada ou saída." />
                </div>
                <div class="collbl">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNota" Style="text-align: right;" runat="server" class="txtNumerico"
                        Width="100px" data-ToolTip="default" ToolTip="Número da Nota Fiscal." />
                </div>
                <div class="collbl">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerie" runat="server" Width="100px" MaxLength="3" data-ToolTip="default"
                        ToolTip="Série da Nota Fiscal." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grdRomaneioNF" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="Produto" HeaderText="Produto">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="30px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Nome" HeaderText="Nome">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="QuantidadeFisica" HeaderText="Físico">
                            <HeaderStyle HorizontalAlign="Left" Width="100px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="100px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="QuantidadeFiscal" HeaderText="Fiscal">
                            <HeaderStyle HorizontalAlign="Left" Width="150px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="150px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Romaneio" HeaderText="Romaneio">
                            <HeaderStyle HorizontalAlign="Left" Width="100px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="100px"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Origem" HeaderText="Origem">
                            <HeaderStyle HorizontalAlign="Left" Width="100px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="100px"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                CHECK
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkRomaneio" runat="server" AutoPostBack="true" OnCheckedChanged="chkRomaneio_CheckedChanged" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
</asp:Content>
