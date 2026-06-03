<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AutorizacoesDePagamentos.aspx.vb" Inherits="NGS.Web.UI.AutorizacoesDePagamentos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngAutorizacoesDePagamentos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAutorizacoesDePagamentos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Autorizações de Pagamentos
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="LinkLiberar" Text="Liberar/Bloquear" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton class="iconPdf" ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
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
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="635px" OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" runat="server" Text="Empresa:"
                        data-ToolTip="default" ToolTip="Consolidar o cnpj da empresa." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="635px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="595px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdCliente" OnClick="cmdCliente_Click" runat="server" Text=">" CssClass="btn"
                        UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" CssClass="txtNumerico" Width="100px" data-ToolTip="default"
                        ToolTip="Número do pedido a ser consultado." />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnPedido" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Número do pedido a ser consultado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" CssClass="calendario" runat="server"
                        Width="100px" data-ToolTip="default" ToolTip="Informar o período inicial de consulta." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" CssClass="calendario" runat="server"
                        Width="100px" CausesValidation="True" data-ToolTip="default" ToolTip="Informar o período final de consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Título:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTitulo" runat="server" CssClass="txtNumerico" Width="100px" data-ToolTip="default"
                        ToolTip="Informar o numero do titulo." />
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:RadioButton ID="rdLiberado" runat="server" GroupName="lib" Text="Liberado" />
                    <asp:RadioButton ID="rdALiberar" runat="server" Checked="True" GroupName="lib" Text="À Liberar" />
                    <asp:RadioButton ID="rdTodos" runat="server" GroupName="lib" Text="Todos" />
                </div>
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Registro
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="coltxt">
                                <asp:Label ID="lblTotalRegistroDolar" runat="server" ForeColor="Red" />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblTotalRegistroReais" runat="server" ForeColor="Red" />
                            </div>
                        </div>
                        <div class="bordagrid" runat="server">
                            <asp:GridView ID="gridRegistro" runat="server" AllowSorting="True" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <ItemStyle Width="30px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Registro" HeaderText="Registro">
                                        <ItemStyle Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                                        HtmlEncode="False">
                                        <ItemStyle Width="80px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Pedido" HeaderText="Pedido">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeProduto" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Historico" HeaderText="Histórico">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Dolar" DataFormatString="{0:N}" HeaderText="Dólares">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Reais" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkTitulos" runat="server" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgSituacao" runat="server" Height="20px" ImageUrl="~/Images/question.JPG"
                                                Width="20px" Style="cursor: pointer; border: 0;" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="UsuarioLiberacao" HeaderText="Autorizante">
                                        <ItemStyle Width="0px"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Produto
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid" runat="server">
                            <asp:GridView ID="gridProduto" runat="server" AllowSorting="True" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" EnableTheming="False"
                                OnSelectedIndexChanged="gridProduto_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <ItemStyle Width="30px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="NomeProduto" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="MoedaLiberado" DataFormatString="{0:N}" HeaderText="U$ Liberado"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="MoedaNaoLiberado" DataFormatString="{0:N}" HeaderText="U$ a Liberar"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorLiberado" DataFormatString="{0:N}" HeaderText="R$ Liberado"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorNaoLiberado" DataFormatString="{0:N}" HeaderText="R$ a Liberar"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkProduto" runat="server" AutoPostBack="True" OnCheckedChanged="chkProduto_CheckedChanged" />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Right" Width="30px" />
                                        <ItemStyle HorizontalAlign="Right" Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel3" runat="server" HeaderText="TabPanel2">
                    <HeaderTemplate>
                        Produto X Registros
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtClienteXProduto" runat="server" Width="300px" />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgBuscarCliente" CssClass="imgconsultar" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                    OnClick="imgBuscarCliente_Click" Enabled="False" />
                            </div>
                            <div class="collbl">
                                Valor a Liberar:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtValor" runat="server" CssClass="txtDecimal" AutoPostBack="True"
                                    OnTextChanged="txtValor_TextChanged" />
                            </div>
                            <div class="collbl">
                                <asp:CheckBox ID="chkProdutoXRegistrosXTodos" runat="server" AutoPostBack="True"
                                    OnCheckedChanged="chkProdutoXRegistrosXTodos_CheckedChanged" Font-Bold="True"
                                    Enabled="False" Text="Marcar Todos" />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblTotalProdutosXRegistros" runat="server" Font-Bold="True" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridProdutoXRegistros" runat="server" AllowSorting="True" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" EnableTheming="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Registro" HeaderText="Registro">
                                        <ItemStyle Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                                        HtmlEncode="False">
                                        <ItemStyle Width="80px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Pedido" HeaderText="Pedido">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeProduto" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Historico" HeaderText="Histórico">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Dolar" DataFormatString="{0:N}" HeaderText="Dólares">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Reais" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkProdutoXRegistros" runat="server" AutoPostBack="True" OnCheckedChanged="chkProdutoXRegistros_CheckedChanged" />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Right" Width="30px" ForeColor="Yellow" />
                                        <ItemStyle HorizontalAlign="Right" Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
