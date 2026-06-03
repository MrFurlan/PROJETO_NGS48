<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="CompensacaoDeAdiantamentos.aspx.vb" Inherits="NGS.Web.UI.CompensacaoDeAdiantamentos" %>
<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1180px !important;
        }

        .collbl {
            width: 125px;
        }

        .w100 {
            width: 100px;
        }

        .w120 {
            width: 120px;
        }

        .lbls {
            min-height: 18px;
            display: block;
            white-space: normal;
            line-height: 16px;
            padding-left: 7px;
            text-indent: 0;
            text-transform: capitalize;
            white-space: normal;
            padding: 4px 0px 4px 7px;
        }

        .txtDecimal {
        }
    </style>
</asp:Content>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngCompensacaoDeAdiantamentos" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="50000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlWFTitulo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hidDataBaixaAdiantamento" runat="server" />
            <div class="titulodiv">
                Compensação de Adiantamentos
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkBaixar" runat="server"
                                Text="Baixar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server"
                                Text="Consultar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server"
                                Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server"
                                Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="595px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" CssClass="btn" runat="server" OnClick="btnCliente_Click"
                        Text=">" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Width="100px" Enabled="false" CssClass="txtNumerico" data-ToolTip="default" ToolTip="Número do pedido a ser consultado." />
                </div>
                <div class="coltxt">
                    <asp:Button CssClass="btn" ID="cmdBuscaPedido" OnClick="cmdBuscaPedido_Click" runat="server"
                        Text="&gt;" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Número do pedido a ser consultado." />
                </div>
                <div class="collbl">
                    Pedido Efetivo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedidoEfetivo" runat="server" Width="100px" Enabled="false" data-ToolTip="default" ToolTip="Número do pedido efetivo a ser consultado." />
                </div>
                <div class="coltxt">
                    <asp:Button CssClass="btn" ID="cmdBuscaPedidoEfetivo" OnClick="cmdBuscaPedidoEfetivo_Click" runat="server"
                        Text="&gt;" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Número do pedido efetivo a ser consultado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Finalidade Financeira:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlFinalidadeFinanceira" runat="server" Width="620px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Adiantamento:
                </div>
                <div class="coltxt">
                    <asp:Button ID="bntAdiantamento" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Títulos:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdPagar" runat="server" AutoPostBack="True" OnCheckedChanged="rdPagar_CheckedChanged" Text="A Pagar" Checked="True"
                        GroupName="titulo" data-ToolTip="default" ToolTip="Selecionar o tipo de títulos a serem consultados." />
                    <asp:RadioButton ID="rdReceber" runat="server" AutoPostBack="True" OnCheckedChanged="rdReceber_CheckedChanged" Text="A Receber" GroupName="titulo"
                        data-ToolTip="default" ToolTip="Selecionar o tipo de títulos a serem consultados." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMoeda" runat="server" Width="127px" AutoPostBack="True" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlIndexador" runat="server" Width="170px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="color: red;">
                    <asp:CheckBox ID="chkDataBaixa" runat="server" Text="Usar Data da Baixa:" Checked="false" AutoPostBack="True" OnCheckedChanged="chkDataBaixa_CheckedChanged" data-ToolTip="default" ToolTip="Se marcado vai usar essa data para Compensação dos Adiantamentos, SE NÃO MARCADO vai usar a data do Vencimento do Título." />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataBaixa" runat="server" Enabled="false" Width="116px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Data da Baixa para Compensação dos Adiantamentos." />
                </div>
            </div>
            <div class="row" id="divAdiantamentos" runat="server" visible="false">
                <div class="bordagrid" style="height: 75px;">
                    <asp:GridView ID="gridAdiantamentos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:BoundField DataField="NumeroDoAdto" HeaderText="Código">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Titulo" HeaderText="Titulo" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Pedido" HeaderText="Pedido" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Conta" HeaderText="Conta" />
                            <asp:BoundField DataField="Safra" HeaderText="Safra">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Vencimento" HeaderText="Vencimento" DataFormatString="{0:dd/MM/yyyy}">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Taxa" HeaderText="Taxa" DataFormatString="{0:N2}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Cifrao" HeaderText="Moeda" />
                            <asp:BoundField DataField="Adiantamento" HeaderText="Vlr.Adto" DataFormatString="{0:N2}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Juros" HeaderText="Juros" DataFormatString="{0:N2}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Baixas" HeaderText="Baixas" DataFormatString="{0:N2}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Saldo" HeaderText="Saldo" DataFormatString="{0:N2}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <div class="row" id="rowDolar" runat="server" visible="False">
                <div class="coltxt">
                    <asp:Label ID="lblTotalRegistroAgrupado" runat="server" Font-Bold="True" />
                    <asp:HiddenField ID="txtRealDolar" runat="server" />
                    <asp:HiddenField ID="txtValorTotal" runat="server" />
                    <asp:HiddenField ID="HiddenIndexador" runat="server" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridConsultaTitulos" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <Columns>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                <asp:CheckBox ID="chkAllTitulos" data-ToolTip="default" ToolTip="Seleciona todos os títulos de mesma moeda e indexador."
                                    Text="CK" runat="server" AutoPostBack="True" OnCheckedChanged="chkAllTitulos_CheckedChanged" />
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkGridTitulos" runat="server" AutoPostBack="True" OnCheckedChanged="chkGridTitulos_CheckedChanged" />
                            </ItemTemplate>
                            <HeaderStyle Width="30px" />
                            <ItemStyle Width="30px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="Registro" HeaderText="Título">
                            <ItemStyle Width="60px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                            HtmlEncode="False">
                            <ItemStyle Width="80px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                            <ItemStyle Width="200px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Historico" HeaderText="Histórico">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" Width="250px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Dolar" DataFormatString="{0:N}" HeaderText="Dolares">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorLiquido" DataFormatString="{0:N}" HeaderText="Reais"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Liberado" HeaderText="Autorizante" ShowHeader="False">
                            <HeaderStyle Width="0px" />
                            <ItemStyle Width="0px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Moeda">
                            <HeaderStyle Width="30px" />
                            <ItemStyle Width="30px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Indexador" HeaderText="Ind">
                            <HeaderStyle HorizontalAlign="Right" Width="20px" />
                            <ItemStyle HorizontalAlign="Right" Width="20px" />
                        </asp:BoundField>
                    </Columns>
                    <EditRowStyle BackColor="#999999" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaAdiantamentos ID="ucConsultaAdiantamentos" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
