<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="LivroDiario.aspx.vb" Inherits="NGS.Web.UI.LivroDiario" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="True"
        AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="upnLivroDiario" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Livro Diário Contábil
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    CNPJ:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCNPJ" Enabled="False" runat="server" Width="122px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nome:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNome" Enabled="false" runat="server" Width="554px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnNome" runat="server" Text=" &gt; " CssClass="btn" data-ToolTip="default" ToolTip="Nome da empresa." />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnLimparEmpresa" CssClass="btn" runat="server" Style="cursor: pointer; background-image: url('images/Borracha.jpg');"
                        Height="22" Width="22" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Processo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtProcesso" runat="server" Style="width: 96px" size="10" CssClass="txtNumerico" data-ToolTip="default" ToolTip="" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnProcesso" runat="server" Text="&gt;" CssClass="btn" />
                </div>
                <div class="collbl">
                    Livro:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtLivro" runat="server" CssClass="txtNumerico" />
                </div>
                <div class="collbl">
                    Folha:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFolha" runat="server" Style="" Width="96px" CssClass="txtNumerico" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataIni" Width="96px" runat="server" class="calendario" TabIndex="5"
                        MaxLength="10" data-ToolTip="default" ToolTip="Data inicial para geração do livro diário." />
                </div>
                <div class="collbl" style="margin-left: 9px;">
                    Período final:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtDataFim" Width="96px" class="calendario" TabIndex="6"
                        MaxLength="10" data-ToolTip="default" ToolTip="Data final para geração do livro diário." />
                </div>
            </div>
            <hr />
            <div class="row">
                <div class="collbl" style="width: 190px;">
                    1- Livro Diário Contabil:
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnApuracao" runat="server" CssClass="botao" Text="Confirmar" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbApuracao" runat="server" CssClass="btn" ImageUrl="images/question.jpg" />
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 190px;">
                    2- Termos Abertura/Encerramento:
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnTermo" runat="server" Text="Confirmar" CssClass="botao" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbTermo" runat="server" CssClass="btn" ImageUrl="images/question.jpg" />
                </div>
            </div>
            <div id="divProcesso" style="width: 800px; height: 600px; display: none">
                <div id="divEmpresa" class="divLoading" style="overflow: auto; height: 300px; width: 100%">
                    <asp:GridView ID="gdvProcesso" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gdvProcesso_SelectedIndexChanged">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Left" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:CommandField ButtonType="button" SelectText=" &gt; " ShowSelectButton="True"
                                ItemStyle-HorizontalAlign="Center" ItemStyle-VerticalAlign="Middle" />
                            <asp:BoundField DataField="Processo_Id" HeaderText="Processo" HeaderStyle-HorizontalAlign="Left" />
                            <asp:BoundField DataField="PeriodoInicial" HeaderText="Período inicial" HeaderStyle-HorizontalAlign="Left" />
                            <asp:BoundField DataField="PeriodoFinal" HeaderText="Período final" HeaderStyle-HorizontalAlign="Left" />
                            <asp:BoundField DataField="Livro" HeaderText="Livro" HeaderStyle-HorizontalAlign="Left" />
                            <asp:BoundField DataField="PaginaInicial" HeaderText="Página" HeaderStyle-HorizontalAlign="Left" />
                        </Columns>
                    </asp:GridView>
                </div>
                <div style="float: right; padding-top: 5px">
                    <asp:Button ID="btnFechar" runat="server" ClientIDMode="Static" Text="Fechar" CssClass="botao" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
</asp:Content>
