<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="TrocaDeNota.aspx.vb" Inherits="NGS.Web.UI.TrocaDeNota" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngTrocaDeNota" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlTrocaDeNota" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdnControlePopup" runat="server" />
            <div class="titulodiv">
                Troca de Nota
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="650px"
                        ToolTip="Empresa Para negociação/consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClientes" runat="server" Width="590px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:HiddenField ID="hdnEnderecoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaCliente" runat="server" Text=" > " UseSubmitBehavior="False"
                        CssClass="btn" OnClick="cmdConsultaCliente_Click" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgLimparCliente" runat="server" Height="20px" ImageAlign="AbsMiddle"
                        CssClass="btn" ImageUrl="images/Borracha.jpg" OnClick="imgLimparCliente_Click"
                        data-ToolTip="default" ToolTip="Limpar Cliente" Width="20px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="90px"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="90px"
                        data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="subtitulodiv">
                Saída/Entrada
            </div>
            <div class="bordagrid" style="height: 315px;">
                <asp:GridView ID="gridTrocaDeNota" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField>
                            <HeaderStyle Width="30px" />
                            <ItemStyle Width="30px" />
                            <ItemTemplate>
                                <asp:ImageButton ID="imgConsultar" runat="server" Height="20px" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                    data-ToolTip="default" ToolTip="Vincular Troca" Width="20px" Style="cursor: pointer"
                                    OnClick="imgConsultar_Click" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField HeaderText="Nota" DataField="Nota" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Movimento" DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Pedido" DataField="CodigoPedido" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Peso" DataField="QuantidadeFiscal" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Empresa" DataField="OEmpresa" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="End" DataField="OEndEmpresa" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Cliente" DataField="OCliente" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="End" DataField="OEndCliente" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Nome" DataField="ONomeCliente" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="E/S" DataField="OEntradaSaida" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="S&#233;r" DataField="OSerie" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Nota" DataField="ONota" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField>
                            <HeaderStyle Width="30px" />
                            <ItemStyle Width="30px" />
                            <ItemTemplate>
                                <asp:ImageButton ID="imgExcluir" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                    data-ToolTip="default" ToolTip="Excluir Vínculo" Width="18px" Style="cursor: pointer"
                                    OnClick="imgExcluir_Click" />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
                <asp:HiddenField ID="HgridRowIndexTroca" runat="server" />
            </div>
            <div class="subtitulodiv">
                Vincular Nota Fiscal:
                <asp:Label ID="lblNotaFiscal" runat="server" Text="Label" />
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresaEntrada" runat="server" Width="650px" Enabled="False"
                        ToolTip="Empresa Para negociação/consulta." />
                </div>
                <div class="coltxt" style="float: right;">
                    <asp:Button ID="btnConsultarEntrada" runat="server" CssClass="botao" Text="Consultar"
                        UseSubmitBehavior="False" Enabled="False" OnClick="btnConsultarEntrada_Click" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClientesEntrada" runat="server" Width="613px" Enabled="false" />
                    <asp:HiddenField ID="txtCodigoClienteEntrada" runat="server" />
                    <asp:HiddenField ID="hdnEnderecoClienteEntrada" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaClienteEntrada" runat="server" Text=" > " UseSubmitBehavior="False"
                        CssClass="btn" Enabled="False" OnClick="cmdConsultaClienteEntrada_Click" data-ToolTip="default"
                        ToolTip="Selecionar o cliente desejado." />
                </div>
                <div class="coltxt" style="float: right;">
                    <asp:Button ID="btnVincular" runat="server" CssClass="botao" Text="Vincular" UseSubmitBehavior="False"
                        Enabled="False" OnClick="btnVincular_Click" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicialEntrada" CssClass="calendario" runat="server" Enabled="False"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinalEntrada" runat="server" CssClass="calendario" Enabled="False"
                        data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtNomeDoCliente" runat="server" ForeColor="Blue" data-ToolTip="default"
                        ToolTip="Nome do Cliente." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtboxEntSai" runat="server" ForeColor="Blue" data-ToolTip="default"
                        ToolTip="" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtboxSerie" runat="server" ForeColor="Blue" data-ToolTip="default"
                        ToolTip="Série da Nota Fiscal." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtboxNota" runat="server" ForeColor="Blue" data-ToolTip="default"
                        ToolTip="Número da Nota Fiscal." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaNotaTroca ID="ucConsultaNotaTroca" runat="server" />
</asp:Content>
