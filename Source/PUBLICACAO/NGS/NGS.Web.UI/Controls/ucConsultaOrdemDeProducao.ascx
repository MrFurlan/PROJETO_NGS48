<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaOrdemDeProducao.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaOrdemDeProducao" %>
<script type="text/javascript">
    function pageLoadConsultaOrdemDeProducao() {
        $("#lnkFecharOrdem", "#divConsultaOrdemDeProducao").button();
        $("#txtLote").setMask("lote-producao");
    }

    function buscarOrdemDeProducao(e) {
        if (e.keyCode == 13) {
            $("#lnkConsultarOrdem", "#divConsultaOrdemDeProducao").click();
            return false;
        }
    }

    $(document).ready(function () {
        pageLoadConsultaOrdemDeProducao();
    });

    var prmConsultaOrdemDeProducao = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaOrdemDeProducao.add_endRequest(pageLoadConsultaOrdemDeProducao);
</script>
<div id="divConsultaOrdemDeProducao" class="uc" title="Consulta Ordem de Producao" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaOrdemDeProducao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultarOrdem" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimparOrdem" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkFecharOrdem" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Código:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSequencia" CssClass="txtNumerico" runat="server" data-ToolTip="default"
                        ToolTip="Sequência de Produção." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Número do Lote:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtLote" runat="server" ClientIDMode="Static" data-ToolTip="default"
                        ToolTip="Número do Lote." />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupoProdutoConsultaProducao" runat="server" AutoPostBack="True" Width="596px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProdutosConsultaProducao" runat="server" Width="596px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt" style="width: 124px;">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="90px"
                        data-ToolTip="default" ToolTip="Informar o período inicial de consulta." />
                </div>
                <div class="collbl" style="margin-left: 5px;">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="90px"
                        data-ToolTip="default" ToolTip="Informar o período final de consulta." />
                </div>
            </div>
            <div class="row">
                <div class="bordagrid">
                    <asp:GridView ID="gridOrdemDeProducao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridOrdemDeProducao_SelectedIndexChanged"
                        Width="100%">
                        <EditRowStyle BackColor="#999999" />
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                <HeaderStyle Width="20px" />
                                <ItemStyle Width="20px" />
                            </asp:CommandField>
                            <asp:BoundField DataField="Codigo" HeaderText="Código">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Lote" HeaderText="Lote">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Movimento" HeaderText="Movimento">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Produto" HeaderText="Produto" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                <ItemStyle HorizontalAlign="Left" Width="100px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="NomeProduto" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
