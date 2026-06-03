<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="VinculoDeNotaFiscal.aspx.vb" Inherits="NGS.Web.UI.VinculoDeNotaFiscal" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngVinculoDeNotaFiscal" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlVinculoDeNotaFiscal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdnControlePopup" runat="server" />
            <asp:HiddenField ID="hdLiberar" runat="server" />
            <div class="titulodiv">
                Vincular/Desvincular Nota Fiscal
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkLiberar" Text="Liberar" runat="server" />
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
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Enabled="false" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="612px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:HiddenField ID="hdnEnderecoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultaCliente" runat="server" OnClick="btnConsultaCliente_Click" Text=">" UseSubmitBehavior="False"
                        CssClass="btn" ToolTip="Consultar Cliente" />
                </div>
            </div>
            <div class="row">
                <div class="collbl colw">
                    Tipo Documento:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlTipoDeDocumento" runat="server" Width="472px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkMovimento" runat="server" AutoPostBack="True" Text="Período:" Checked="false" data-ToolTip="default" ToolTip="Marcar se desejar inserir data para consulta." />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="100px" />
                    &nbsp;
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="100px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nota Fiscal:
                </div>
                <div class="coltxt" style="width: 135px;">
                    <asp:TextBox ID="txtNotaFiscal" CssClass="txtNumerico" runat="server" Width="100px" />
                </div>
                <div id="divES" runat="server">
                    <div class="collbl">
                        Tipo:
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rdEntrada" runat="server" Text="Entrada" GroupName="nfentsai" />
                        <asp:RadioButton ID="rdSaida" runat="server" Text="Saída" GroupName="nfentsai" />
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="bordagrid" style="height: 415px;">
                    <asp:GridView ID="gridVinculoDeNota" runat="server" AutoGenerateColumns="false" CellPadding="4"
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
                                        ToolTip="Vincular Nota" Width="20px" Style="cursor: pointer" OnClick="imgConsultar_Click" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField HeaderText="Empresa" DataField="Empresa" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField HeaderText="Cliente" DataField="Cliente" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
<%--                            <asp:BoundField HeaderText="Nome" DataField="NomeCliente" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" Width="200px" />
                                <ItemStyle HorizontalAlign="Left" Width="200px" />
                            </asp:BoundField>--%>
                            <asp:BoundField HeaderText="E/S" DataField="EntradaSaida_Id" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField HeaderText="S&#233;r" DataField="Serie_Id" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField HeaderText="Nota" DataField="Nota_Id" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}" DataField="Movimento" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField HeaderText="Descrição" DataField="Descricao" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" Width="200px" />
                                <ItemStyle HorizontalAlign="Left" Width="200px" />
                            </asp:BoundField>
                            <asp:TemplateField>
                                <HeaderStyle Width="30px" />
                                <ItemStyle Width="30px" />
                                <ItemTemplate>
                                    <asp:ImageButton ID="imgExcluir" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                        ToolTip="Excluir Vínculo" Width="18px" Style="cursor: pointer" OnClick="imgExcluir_Click" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField HeaderText="Empresa" DataField="OrigemEmpresa" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField HeaderText="Cliente" DataField="OrigemCliente" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
<%--                            <asp:BoundField HeaderText="Nome" DataField="ONomeCliente" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" Width="200px" />
                                <ItemStyle HorizontalAlign="Left" Width="200px" />
                            </asp:BoundField>--%>
                            <asp:BoundField HeaderText="E/S" DataField="OrigemEntradaSaida_Id" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField HeaderText="S&#233;r" DataField="OrigemSerie_Id" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField HeaderText="Nota" DataField="OrigemNota_id" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}" DataField="OrigemMovimento" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField HeaderText="Descrição" DataField="OrigemDescricao" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" Width="200px" />
                                <ItemStyle HorizontalAlign="Left" Width="200px" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                    <asp:HiddenField ID="HgridRowIndexVinculo" runat="server" />
                </div>
            </div>
            <div class="row" runat="server">
                <div id="pnlVinculoNota" class="painelleft" runat="server" style="width: 100%;">
                    <div class="subtitulodiv">
                        Vincular Nota Fiscal:
                    <asp:Label ID="lblNotaFiscal" runat="server" Text="Label" />
                    </div>
                    <div class="menu_acoes" runat="server">
                        <div class="acoes" runat="server">
                            <ul>
                                <li class="iconConsultar" runat="server">
                                    <asp:LinkButton ID="lnkConsultarVinc" Text="Consultar" runat="server" />
                                </li>
                                <li class="iconRelatorio" runat="server">
                                    <asp:LinkButton ID="lnkVincular" Text="Vincular" runat="server" />
                                </li>
                            </ul>
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Empresa:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlEmpresaVinc" runat="server" Width="650px" Enabled="false" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Cliente:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtClientesVinc" runat="server" Width="610px" Enabled="False" />
                            <asp:HiddenField ID="txtCodigoClienteVinc" runat="server" />
                            <asp:HiddenField ID="hdnEnderecoClienteVinc" runat="server" />
                        </div>
                        <div class="coltxt">
                            <asp:Button ID="btnConsultaClienteVinc" runat="server" Text=">" UseSubmitBehavior="False"
                                CssClass="btn" OnClick="btnConsultaClienteVinc_Click" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Data Inicial:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtDataInicialVinc" CssClass="calendario" runat="server" />
                        </div>
                        <div class="collbl">
                            Data Final:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtDataFinalVinc" runat="server" CssClass="calendario" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Nota Fiscal:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtNotaVinc" CssClass="txtNumerico" runat="server" Width="100px" />
                        </div>
                        <div class="collbl" style="margin-left: 21px;">
                            Tipo:
                        </div>
                        <div class="coltxt">
                            <asp:RadioButton ID="rdEntradaVinc" runat="server" Text="Entrada" GroupName="nfentsaiVinc" />
                            <asp:RadioButton ID="rdSaidaVinc" runat="server" Text="Saída" GroupName="nfentsaiVinc" />
                        </div>
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaNotaTroca ID="ucConsultaNotaTroca" runat="server" />
    <uc:consultaliberacao id="ucLiberacao" runat="server" />
</asp:Content>
