<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CarteiraDeCompras.aspx.vb" Inherits="NGS.Web.UI.CarteiraDeCompras" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCarteiraDeCompras" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCarteiraDeCompras" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdnContaSelecionada" runat="server" />
            <div class="titulodiv">
                Finalidade Financeira
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
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
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Código:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" runat="server" MaxLength="9" Width="120px" data-ToolTip="default"
                        ToolTip="Código de cadastro da carteira de compras." />
                </div>
                <div class="collbl" style="margin-left: 103px;">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSituacao" runat="server" Width="125px" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkPedido" runat="server" Text="Pedido" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" Width="609px" data-ToolTip="default"
                        ToolTip="Descrição da carteira de compras." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classificação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlClassificacao" runat="server" Width="233px" />
                </div>
                <div class="collbl">
                    Adiantamento:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="radSim" runat="server" Text="Sim" GroupName="gpnAdiantamento"
                        data-ToolTip="default" ToolTip="Selecionar se é adiantamento ou não." AutoPostBack="true" />
                    <asp:RadioButton ID="radNao" runat="server" Text="Não" GroupName="gpnAdiantamento"
                        data-ToolTip="default" ToolTip="Selecionar se é adiantamento ou não." AutoPostBack="true" />
                </div>

                <div class="coltxt" id="divBaixaDeAdiantamento" runat="server">
                    <asp:CheckBox ID="chkBaixaDeAdiantamento" runat="server" Text="Baixa de Adiantamento" AutoPostBack="true" />
                </div>
            </div>

            <div class="row">
                <div class="collbl">
                    Conta Clientes:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaCliente" runat="server" Enabled="false" MaxLength="50" OnFocus="return false"
                        ReadOnly="true" Width="560px" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbContaCliente" runat="server" align="absMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                        OnClick="imbContaCliente_Click" data-ToolTip="default" ToolTip="Consultar Clientes"
                        Style="margin-top: 5px;" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbLimparContaCliente" runat="server" align="absMiddle" Height="22"
                        CssClass="btn" src="images/Borracha.jpg" Style="cursor: pointer" data-ToolTip="default"
                        ToolTip="Limpar Clientes" Width="22" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta Descontos:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaDesconto" runat="server" Enabled="false" MaxLength="50"
                        ReadOnly="true" Width="560px" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbPesquisarDesconto" runat="server" align="absMiddle" Style="margin-top: 5px;"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imbPesquisarDesconto_Click"
                        data-ToolTip="default" ToolTip="Conta para lançamento quando houver descontos." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbLimparDesconto" runat="server" align="absMiddle" Height="22"
                        CssClass="btn" src="images/Borracha.jpg" Style="cursor: pointer" data-ToolTip="default"
                        ToolTip="Limpar Clientes" Width="22" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta Deduções:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaDeducoes" runat="server" Enabled="false" MaxLength="50"
                        ReadOnly="true" Width="560px" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbPesquisarDeducoes" runat="server" align="absMiddle" Style="margin-top: 5px;"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imbPesquisarDeducoes_Click"
                        data-ToolTip="default" ToolTip="Conta para lançamento quando houver deduções." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbLimparDeducooes" runat="server" align="absMiddle" Height="22"
                        src="images/Borracha.jpg" CssClass="btn" data-ToolTip="default" ToolTip="Limpar Clientes"
                        Width="22" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta Juros:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaJuros" runat="server" Enabled="false" MaxLength="50" ReadOnly="true"
                        Width="560px" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbPesquisarJuros" runat="server" align="absMiddle" Style="margin-top: 5px;"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imbPesquisarJuros_Click"
                        data-ToolTip="default" ToolTip="Conta para lançamento quando houver juros." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbLimparJuros" runat="server" align="absMiddle" Height="22"
                        CssClass="btn" src="images/Borracha.jpg" Style="cursor: pointer" data-ToolTip="default"
                        ToolTip="Limpar Clientes" Width="22" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta Acréscismos:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaAcrescimos" runat="server" Enabled="false" MaxLength="50"
                        ReadOnly="true" Width="560px" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbPesquisarAcrescimos" runat="server" align="absMiddle" Style="margin-top: 5px;"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imbPesquisarAcrescimos_Click"
                        data-ToolTip="default" ToolTip="Conta para lançamento quando houver acréscimos." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbLimparAcrescimos" runat="server" align="absMiddle" Height="22"
                        CssClass="btn" src="images/Borracha.jpg" Style="cursor: pointer" data-ToolTip="default"
                        ToolTip="Limpar Clientes" Width="22" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gdvCarteiraDeCompras" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gdvCarteiraDeCompras_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" ItemStyle-HorizontalAlign="Center" ItemStyle-VerticalAlign="Middle"
                            SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Produto_Id" HeaderText="Código" />
                        <asp:BoundField DataField="Descricao" HeaderText="Descrição" />
                        <asp:BoundField DataField="Situacao" HeaderText="Sit." />
                        <asp:BoundField DataField="ContaClientes" HeaderText="Clientes" />
                        <asp:BoundField DataField="ContaDescontos" HeaderText="Descontos" />
                        <asp:BoundField DataField="ContaDeducoes" HeaderText="Deduções" />
                        <asp:BoundField DataField="ContaJuros" HeaderText="Juros" />
                        <asp:BoundField DataField="ContaAcrescimos" HeaderText="Acréscimos" />
                        <asp:BoundField DataField="Classificacao" HeaderText="Clas." />
                        <asp:BoundField DataField="NomeContaCliente" HeaderStyle-CssClass="none" ItemStyle-CssClass="none" />
                        <asp:BoundField DataField="NomeContaDesconto" HeaderStyle-CssClass="none" ItemStyle-CssClass="none" />
                        <asp:BoundField DataField="NomeContaDeducao" HeaderStyle-CssClass="none" ItemStyle-CssClass="none" />
                        <asp:BoundField DataField="NomeContaJuro" HeaderStyle-CssClass="none" ItemStyle-CssClass="none" />
                        <asp:BoundField DataField="NomeContaAcrescimo" HeaderStyle-CssClass="none" ItemStyle-CssClass="none" />
                        <asp:BoundField DataField="Adiantamento" HeaderText="Adto." />
                        <asp:BoundField DataField="Pedido" HeaderText="Pedido" />
                        <asp:BoundField DataField="BaixaAdiantamento" HeaderText="Baixa de Adto." />
                        <asp:BoundField DataField="CarteiraBaixaAdiantamento" HeaderText="Carteira de Baixa de Adto." />
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
    <uc:ConsultaContaCliente ID="ucConsultaContaCliente" runat="server" />
</asp:Content>
