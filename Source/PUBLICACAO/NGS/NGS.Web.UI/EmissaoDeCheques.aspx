<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="EmissaoDeCheques.aspx.vb" Inherits="NGS.Web.UI.EmissaoDeCheques" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngEmissaoDeCheques" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlEmissaoDeCheques" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Emissão/Reemissão de Cheques
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>

                        <li class="iconRelatorio rel" runat="server"><a>Emitir</a>
                            <ul>
                                <li class="iconRelatorio">
                                    <asp:LinkButton ID="lnkEmitir" runat="server" Text="Padrão" />
                                </li>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkEmitirPdf" runat="server" Text="Pdf" />
                                </li>
                            </ul>
                        </li>

                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" runat="server" Text="Relatório" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlUnidadeDeNegocioEmpresaCliente" runat="server" Width="590px"
                        AutoPostBack="True" OnSelectedIndexChanged="DdlUnidadeDeNegocioEmpresaCliente_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa Pagadora:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresaCliente" runat="server" Width="590px" AutoPostBack="true"
                        OnSelectedIndexChanged="DdlEmpresaCliente_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Banco:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlBancoPagador" runat="server" Width="590px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlBancoPagador_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlContaPagadora" TabIndex="3" runat="server" Width="590px"
                        AutoPostBack="True" OnSelectedIndexChanged="DdlContaPagadora_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt" style="width: 118px;">
                    <asp:TextBox ID="txtMovimento" CssClass="calendario" runat="server" Width="85px"
                        data-ToolTip="default" ToolTip="Inserir a data que está sendo realizada o lançamento." />
                </div>
                <div class="collbl" style="width: 128px;">
                    <asp:CheckBox ID="chkReimprime" runat="server" Text="Reimprime Cheque" data-ToolTip="default"
                        ToolTip="Selecionar para reimprimir o cheque." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Num Cheque:
                </div>
                <div class="coltxt" style="width: 118px;">
                    <asp:TextBox ID="txtChequeinicial" runat="server" Width="85px" CssClass="txtNumerico"
                        Enabled="false" />
                </div>
                <div class="collbl" style="width: 128px;">
                    Impressoras:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlImpressora" runat="server" Width="328px" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridConsultaTitulos1" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="Registro" SortExpression="Registro" HeaderText="Registro"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="80px" />
                            <ItemStyle HorizontalAlign="Left" Width="80px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Sequencia" SortExpression="Sequencia" HeaderText="Sequ&#234;ncia"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="50px" />
                            <ItemStyle HorizontalAlign="Left" Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Lote" SortExpression="Lote" HeaderText="Lote" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="50px" />
                            <ItemStyle HorizontalAlign="Left" Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Destinatario" HeaderText="Destinat&#225;rio" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="200px" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Historico" SortExpression="Historico" HeaderText="Hist&#243;rico"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Valor" HeaderText="Valor" HtmlEncode="False">
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                            <HeaderStyle HorizontalAlign="Right" Width="100px" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Imprime?">
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemTemplate>
                                <asp:CheckBox ID="ChkImprimir" runat="server" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Cidade" HeaderText="Cidade Pgto" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="200px" />
                            <ItemStyle HorizontalAlign="Left" Width="200px" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Destina&#231;&#227;o">
                            <ItemTemplate>
                                <asp:TextBox ID="txtDestinacao" runat="server" Enabled="False" Text='<%# eval("Destinacao") %>'
                                    Width="200px" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="ContaContabilPagadora" HeaderText="Conta Cont&#225;bil"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="80px" />
                            <ItemStyle HorizontalAlign="Left" Width="80px" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
