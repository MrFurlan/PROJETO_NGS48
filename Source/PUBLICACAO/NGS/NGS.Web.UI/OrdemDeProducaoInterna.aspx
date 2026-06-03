<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="OrdemDeProducaoInterna.aspx.vb" Inherits="NGS.Web.UI.OrdemDeProducaoInterna" %>
<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }
    </style>
    <script type="text/javascript">
        function pageLoadOrdemDeProducaoInterna() {
            //$("#txtLoteInterno").setMask("lote-producao");
        }

        $(document).ready(function () {
            pageLoadOrdemDeProducaoInterna();
            var prmItemProducao = Sys.WebForms.PageRequestManager.getInstance();
            prmItemProducao.add_endRequest(pageLoadOrdemDeProducaoInterna);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngOrdemDeProducaoInterna" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlOrdemDeProducaoInterna" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Ordem de Produção Interna
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovoI" Text="Gravar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizarI" Text="Atualizar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultarI" Text="Consultar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimparI" Text="Limpar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" Enabled="False" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeNegocio_SelectedIndexChanged" Width="596px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Enabled="False" AutoPostBack="True" OnSelectedIndexChanged="ddlEmpresa_SelectedIndexChanged" Width="596px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupoProdutoInterno" runat="server" Enabled="False" AutoPostBack="True" Width="576px" />
                </div>
                <div class="coltxt">
                    <asp:LinkButton ID="lnkBuscaProdutoInterno" runat="server" Enabled="False" Height="20px" Width="20px"><asp:Image ID="Image1" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                            ToolTip="Consulta Produto" /></asp:LinkButton>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProdutoInterno" runat="server" Enabled="False" AutoPostBack="True" Width="596px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Entrada/Saída:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdEntradas" runat="server" Text="Entradas" GroupName="OrdemInterna" AutoPostBack="True" data-ToolTip="default" ToolTip="Informar se é entrada ou saída." />
                    <asp:RadioButton ID="rdSaidas" runat="server" Text="Saidas" GroupName="OrdemInterna" AutoPostBack="True" data-ToolTip="default" ToolTip="Informar se é entrada ou saída." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbOperacao" Enabled="false" runat="server" Width="296px" />
                    <asp:DropDownList ID="cmbSubOperacao" Enabled="false" runat="server" Width="296px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Número do Lote:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtLoteInterno" runat="server" ClientIDMode="Static" MaxLength="20" Enabled="False" data-ToolTip="default"
                        ToolTip="Número do Lote no máximo 20 caractéres." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data de Produção:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataProducaoInterna" CssClass="calendario" runat="server" Enabled="False" data-ToolTip="default"
                        ToolTip="Data inicial do lançamento." />
                </div>
                <div class="collbl">
                    Data de Validade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataValidadeInterna" CssClass="calendario" runat="server" Enabled="False" data-ToolTip="default"
                        ToolTip="Data de Validade." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Quantidade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtQuantidadeInterna" runat="server" Enabled="False" CssClass="txtDecimal4" data-ToolTip="default" ToolTip="Quantidade de Produção." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Observações:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtObsProdutoInterno" runat="server" Width="750px" Height="50px" TextMode="MultiLine"
                        data-ToolTip="default" ToolTip="Preencher quando houver alguma observação relevante." />
                </div>
            </div>
            <div class="row">
                <div class="bordagrid" style="width: 100%; height: 315px;">
                    <asp:GridView ID="gridProducaoInterno" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                        <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:TemplateField ShowHeader="False">
                                <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                <ItemStyle HorizontalAlign="Center" Width="30px" />
                                <ItemTemplate>
                                    <asp:ImageButton ID="imgSelecionaPrdInterno" runat="server" ImageUrl="~/images/select.jpg"
                                        OnClick="imgSelecionaPrdInterno_Click" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="Produto" HeaderText="Produto" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                <ItemStyle HorizontalAlign="Left" Width="100px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="NomeProduto" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Movimento" HeaderText="Data da Produção" DataFormatString="{0:dd/MM/yyyy}" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Validade" HeaderText="Validade" DataFormatString="{0:dd/MM/yyyy}" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Lote" HeaderText="Lote" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Observacoes" HeaderText="Observações" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="EntradaSaida" HeaderText="E/S" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Operacao" HeaderText="OP" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="SubOperacao" HeaderText="SO" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:ImageButton ID="imgRemoverProdutoInterno" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                        OnClick="imgRemoverProdutoInterno_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente remover o Produto?');" />
                                </ItemTemplate>
                                <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
</asp:Content>
