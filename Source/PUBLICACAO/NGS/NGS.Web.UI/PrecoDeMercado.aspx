<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PrecoDeMercado.aspx.vb" Inherits="NGS.Web.UI.PrecoDeMercado" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPrecoDeMercado" runat="server" EnablePartialRendering="true"
        EnableScriptGlobalization="True" EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPrecoDeMercado" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Preço de Mercado
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton ID="lnkPdf" class="iconPdf" runat="server" Text="Pdf" />
                                </li>
                                <li>
                                    <asp:LinkButton ID="lnkExcel" class="iconExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
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
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Deposito:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlDeposito" runat="server" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupo" runat="server" OnSelectedIndexChanged="ddlGrupo_SelectedIndexChanged1"
                        Width="650px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProduto" runat="server" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt" style="width: 117px;">
                    <asp:TextBox ID="txtData" Width="85px" CssClass="calendario" runat="server" data-ToolTip="default"
                        ToolTip="Data do lançamento." />
                </div>
                <div class="collbl">
                    Valor Oficial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValorOficial" runat="server" AutoPostBack="True" OnTextChanged="txtValorOficial_TextChanged"
                        CssClass="txtDecimal6" data-ToolTip="default" ToolTip="Valor do produto na data atual." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Valor Moeda:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValorMoeda" runat="server" AutoPostBack="True" OnTextChanged="txtValorMoeda_TextChanged"
                        CssClass="txtDecimal6" data-ToolTip="default" ToolTip="Valor da cotação diária." />
                </div>
                <div class="collbl" style="margin-left: 7px;">
                    Base De Calculo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtBaseDeCalculo" runat="server" MaxLength="10" CssClass="txtNumerico"
                        data-ToolTip="default" ToolTip="A quantidade do produto comercializado." />
                </div>
            </div>
            <div class="subtitulodiv">
                Consulta
            </div>
            <div class="row">
                <div class="collbl">
                    Data de:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDatade" CssClass="calendario" runat="server" Width="86px" data-ToolTip="default"
                        ToolTip="Data inicial de consulta" />
                </div>
                <div class="collbl">
                    até:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataAte" CssClass="calendario" runat="server" data-ToolTip="default"
                        ToolTip="Data final de consulta" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridPrecoDeMercado" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridPrecoDeMercado_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                            <ItemStyle Width="30px" />
                        </asp:CommandField>
                        <asp:BoundField DataField="EmpresaNome" HeaderText="Empresa" HtmlEncode="False" />
                        <asp:BoundField DataField="DepositoNome" HeaderText="Deposito" HtmlEncode="False" />
                        <asp:BoundField DataField="ProdutoNome" HeaderText="Produto" HtmlEncode="False"></asp:BoundField>
                        <asp:BoundField DataField="Data_Id" HeaderText="Data" DataFormatString="{0:d}" HtmlEncode="False"></asp:BoundField>
                        <asp:BoundField DataField="ValorOficial" HeaderText="Valor Oficial" HtmlEncode="False">
                            <ItemStyle HorizontalAlign="Right" />
                            <HeaderStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorMoeda" HeaderText="Valor Moeda" HtmlEncode="False">
                            <ItemStyle HorizontalAlign="Right" />
                            <HeaderStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="BaseDeCalculo" HeaderText="Base De Calculo" HtmlEncode="False">
                            <ItemStyle HorizontalAlign="Right" />
                            <HeaderStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
