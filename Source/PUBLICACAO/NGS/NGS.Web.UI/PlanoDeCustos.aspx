<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PlanoDeCustos.aspx.vb" Inherits="NGS.Web.UI.PlanoDeCustos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 125px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPlanoDeCustos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPlanoDeCustos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Plano De Custos
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Lançamento
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
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
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigo" runat="server" Width="48px" MaxLength="4" data-ToolTip="default"
                                    ToolTip="Referente ao registro da conta de custos." />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricao" runat="server" Width="457px" MaxLength="99" data-ToolTip="default"
                                    ToolTip="Referente ao registro da conta de custos." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Totalizador:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTotalizador" runat="server" Width="529px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Sinal:
                            </div>
                            <div class="coltxt">
                                Peso :&nbsp;
                                <asp:DropDownList ID="DdlSinalPeso" runat="server" />
                            </div>
                            <div class="coltxt">
                                Valor :
                                <asp:DropDownList ID="DdlSinalValor" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="ChkDesdobramento" runat="server" Text="Desdobramento" data-ToolTip="default"
                                    ToolTip="Se o sinal do lançamento será positivo ou negativo no peso e no valor." />
                                <asp:CheckBox ID="ChkSaldoInicial" runat="server" Text="Saldo Inicial" data-ToolTip="default"
                                    ToolTip="Se o sinal do lançamento será positivo ou negativo no peso e no valor." />
                                <asp:CheckBox ID="chkRateio" runat="server" Text="Rateio" data-ToolTip="default"
                                    ToolTip="Se o sinal do lançamento será positivo ou negativo no peso e no valor." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Débito Mercadoria:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlDebitaMercadoria" runat="server" Width="529px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Crédito Mercadoria:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlCreditaMercadoria" runat="server" Width="529px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Histórico Mercadoria:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlHisMercadoria" runat="server" Width="529px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Débito Fretes:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlDebitaFrete" runat="server" Width="529px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Crédito Fretes:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlCreditaFrete" runat="server" Width="529px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Histórico Fretes:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlHisFretes" runat="server" Width="529px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conta Origem:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlContaOrigem" runat="server" Width="438px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto Origem:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlProdutoOrigem" runat="server" Width="438px" />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="BtnAdiciona" OnClick="BtnAdiciona_Click" runat="server" ImageUrl="~/images/ico-mais.gif"
                                    CssClass="btn" ImageAlign="AbsMiddle" data-ToolTip="default" ToolTip="Adicionar item na Lista" />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="btnRemove" OnClick="btnRemove_Click" runat="server" ImageUrl="~/images/ico-menos.gif"
                                    CssClass="btn" ImageAlign="AbsMiddle" data-ToolTip="default" ToolTip="Remover item da Lista" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 150px;">
                            <asp:GridView ID="gridOrigem" runat="server" AutoGenerateColumns="False" Width="100%"
                                CellPadding="1" GridLines="None" ForeColor="#333333" OnSelectedIndexChanged="gridOrigem_SelectedIndexChanged">
                                <EditRowStyle BackColor="#999999" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Conta_Id" HeaderText="Cód. Conta">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Titulo" HeaderText="Dsc. Conta">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Produto_Id" HeaderText="Cód. Produto">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Produto" HeaderText="Dsc. Produto">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="GridPlanoDeCustos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridView1_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <ItemStyle Width="25px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center" Width="50px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Totalizador" HeaderText="Tot"></asp:BoundField>
                                    <asp:BoundField DataField="SinalPeso" HeaderText="SP"></asp:BoundField>
                                    <asp:BoundField DataField="SinalValor" HeaderText="SV"></asp:BoundField>
                                    <asp:BoundField DataField="DebitoMercadoria" HeaderText="Deb.Merc."></asp:BoundField>
                                    <asp:BoundField DataField="CreditoMercadoria" HeaderText="Cred.Merc."></asp:BoundField>
                                    <asp:BoundField DataField="HistoricoMercadoria" HeaderText="Hst"></asp:BoundField>
                                    <asp:BoundField DataField="DebitoFrete" HeaderText="Deb.Frete"></asp:BoundField>
                                    <asp:BoundField DataField="CreditoFrete" HeaderText="Cred.Frete"></asp:BoundField>
                                    <asp:BoundField DataField="HistoricoFrete" HeaderText="Hst"></asp:BoundField>
                                    <asp:CheckBoxField DataField="SaldoInicial" HeaderText="S.I." ItemStyle-HorizontalAlign="Center"
                                        ItemStyle-VerticalAlign="Middle"></asp:CheckBoxField>
                                    <asp:CheckBoxField DataField="Desdobramento" HeaderText="Desd." ItemStyle-HorizontalAlign="Center"
                                        ItemStyle-VerticalAlign="Middle"></asp:CheckBoxField>
                                    <asp:CheckBoxField DataField="Rateio" HeaderText="Rat." ItemStyle-HorizontalAlign="Center"
                                        ItemStyle-VerticalAlign="Middle"></asp:CheckBoxField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
