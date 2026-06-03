<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ProdutosXPrecos.aspx.vb" Inherits="NGS.Web.UI.ProdutosXPrecos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngProdutosXPrecos" runat="server" EnablePartialRendering="true"
        EnableScriptGlobalization="True" EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlProdutosXPrecos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Produtos X Preços
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Manutenção
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                                    </li>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
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
                                Tabela de Preço:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTabelaDePreco" runat="server" Width="650px" AutoPostBack="True" OnSelectedIndexChanged="ddlTabelaDePreco_SelectedIndexChanged" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCliente" runat="server" Width="640px" Enabled="False" />
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button CssClass="btn" ID="btnBuscaCliente" OnClick="btnBuscaCliente_Click" runat="server"
                                    Text=">" CausesValidation="False" UseSubmitBehavior="False" data-tooltip="default"
                                    ToolTip="Selecionar o cliente desejado." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupo" runat="server" Width="650px" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlGrupo_SelectedIndexChanged" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProduto" runat="server" Width="650px" AutoPostBack="True" OnSelectedIndexChanged="ddlProduto_SelectedIndexChanged" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnProduto" runat="server" OnClick="btnProduto_Click" Text=" > "
                                    CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Forma de compra/venda do produto." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMovimento" CssClass="calendario" runat="server" data-ToolTip="default" ToolTip="Data do registro do produto." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Moeda:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlMoeda" runat="server" Width="123px" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlMoeda_SelectedIndexChanged" />
                            </div>
                            <div class="collbl">
                                Valor:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="TxtValor" CssClass="txtDecimal" runat="server" OnTextChanged="TxtValor_TextChanged"
                                    AutoPostBack="True" data-ToolTip="default" ToolTip="Valor em real do produto." />
                            </div>
                            <div class="collbl" style="margin-left: 52px;">
                                Dolar:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDolar" CssClass="txtDecimal" runat="server" AutoPostBack="True" OnTextChanged="txtDolar_TextChanged"
                                    data-ToolTip="default" ToolTip="Valor em dolar do produto." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Fixo Operacional:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFixoOperacional" CssClass="txtDecimal" runat="server" AutoPostBack="True"
                                    OnTextChanged="txtFixoOperacional_TextChanged" Width="49px" data-ToolTip="default" ToolTip="" />
                                <asp:Label ID="Label4" runat="server" Font-Bold="True" Text=" % " />
                            </div>
                            <div class="collbl" style="margin-left: 45px;">
                                Margem Menor:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMargemMenor" CssClass="txtDecimal" runat="server" Width="50px"
                                    AutoPostBack="True" OnTextChanged="txtMargemMenor_TextChanged" data-ToolTip="default" ToolTip="" />
                                <asp:Label ID="Label1" runat="server" Font-Bold="True" Text=" % " />
                                <asp:TextBox ID="txtValorMargemMenor" CssClass="txtDecimal" runat="server" Width="69px" data-ToolTip="default" ToolTip="" />
                            </div>
                            <div class="collbl">
                                Margem Maior:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMargemMaior" CssClass="txtDecimal" runat="server" Width="50px"
                                    AutoPostBack="True" OnTextChanged="txtMargemMaior_TextChanged" data-ToolTip="default" ToolTip="" />
                                <asp:Label ID="Label3" runat="server" Font-Bold="True" Text=" % " />
                                <asp:TextBox ID="txtValorMargemMaior" CssClass="txtDecimal" runat="server" Width="70px" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>

                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Tabela de Preço
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaTabela" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtData" CssClass="calendario" runat="server" AutoPostBack="True"
                                    OnTextChanged="txtDataTabela_TextChanged" data-ToolTip="default" ToolTip="Data do registro do produto." />
                            </div>
                            <div class="collbl">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupoPreco" runat="server" Width="215px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 650px;">
                            <asp:GridView ID="gridPxP" runat="server" ForeColor="#333333" GridLines="None" Width="100%"
                                CellPadding="4" AutoGenerateColumns="False" OnSelectedIndexChanged="gridPxP_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#5D7B9D" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                                    <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeProduto" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Wrap="False"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Movimento" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeMoeda" HeaderText="Moeda">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" HeaderText="Valor" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="MargemMenor" HeaderText=" &lt;">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorMargemMenor" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="MargemMaior" HeaderText="&gt;">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorMargemMaior" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>

                <ajaxToolkit:TabPanel ID="TabPanel3" runat="server">
                    <HeaderTemplate>
                        Histórico Preço Produto
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="subtitulodiv">
                            <asp:Label ID="lblHistoricoProduto" runat="server" />
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridHistorico" runat="server" Width="100%" GridLines="None" ForeColor="#333333"
                                CellPadding="4" AutoGenerateColumns="False" OnSelectedIndexChanged="gridHistorico_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                                    <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeProduto" HeaderText="Nome" Visible="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Movimento" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeMoeda" HeaderText="Moeda">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" HeaderText="Valor" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="MargemMenor" HeaderText=" &lt;">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorMargemMenor" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="MargemMaior" HeaderText="&gt;">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorMargemMaior" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
