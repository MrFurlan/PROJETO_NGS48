<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaPlanoDeContas.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaPlanoDeContas" %>
<script type="text/javascript">
    function pageLoadConsultaPlanoDeContas() {
        $("#btnFechar", "#divConsultaPlanoDeContas").button();
        $("#btnLimpar", "#divConsultaPlanoDeContas").button();
        $("#btnConsultarConta", "#divConsultaPlanoDeContas").button();
        $("#btnConsultarSubConta", "#divConsultaPlanoDeContas").button();
        $("#txtConta", "#divConsultaPlanoDeContas").focus();
    }

    $(document).ready(function () {
        pageLoadConsultaPlanoDeContas();
    });

    var prmConsultaPlanoDeContas = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaPlanoDeContas.add_endRequest(pageLoadConsultaPlanoDeContas);
</script>
<style type="text/css">
    .collbluc
    {
        width: 113px !important;
    }
</style>
<div id="divConsultaPlanoDeContas" class="uc" title="Consulta Plano de Contas" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaPlanoDeContas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:Panel ID="pnlParametros" runat="server" Width="100%">
                <div class="menu_acoes">
                    <div class="acoes">
                        <ul>
                            <li class="iconConsultar" runat="server">
                                <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                            </li>
                            <li class="iconLimpar" runat="server">
                                <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                            </li>
                            <li class="iconSair" runat="server">
                                <asp:LinkButton ID="lnkSair" Text="Fechar" runat="server" />
                            </li>
                        </ul>
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Conta:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtConta" runat="server" CssClass="txtNumerico" ClientIDMode="Static"
                            Width="104px" />
                    </div>
                    <div class="collbluc" style="margin-left: 153px;">
                        Encargo Financeiro:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlTemEncargoFinanceiro" runat="server" Width="60px">
                            <asp:ListItem Value=""></asp:ListItem>
                            <asp:ListItem Value="S">Sim</asp:ListItem>
                            <asp:ListItem Value="N">Não</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                    <div class="collbluc">
                        Tem Pedido:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlTemPedido" runat="server" Width="60px">
                            <asp:ListItem Value=""></asp:ListItem>
                            <asp:ListItem Value="S">Sim</asp:ListItem>
                            <asp:ListItem Value="N">Não</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Titulo:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtTitulo" runat="server" Width="257px" />
                    </div>
                    <div class="collbluc">
                        É Adiantamento:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlAdiantamento" runat="server" Width="60px">
                            <asp:ListItem Value=""></asp:ListItem>
                            <asp:ListItem Value="S">Sim</asp:ListItem>
                            <asp:ListItem Value="N">Não</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                    <div class="collbluc">
                        Tem Produto:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlTemProduto" runat="server" Width="60px">
                            <asp:ListItem Value=""></asp:ListItem>
                            <asp:ListItem Value="S">Sim</asp:ListItem>
                            <asp:ListItem Value="N">Não</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Tipo Conta:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlTipoDeConta" runat="server" Width="267px" />
                    </div>
                    <div class="collbluc">
                        Tem Cliente:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlTemCliente" runat="server" Width="60px">
                            <asp:ListItem Value=""></asp:ListItem>
                            <asp:ListItem Value="S">Sim</asp:ListItem>
                            <asp:ListItem Value="N">Não</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                    <div class="collbluc">
                        Tem Centro Custo:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlTemCentroCusto" runat="server" Width="60px">
                            <asp:ListItem Value=""></asp:ListItem>
                            <asp:ListItem Value="S">Sim</asp:ListItem>
                            <asp:ListItem Value="N">Não</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
            </asp:Panel>
            <div class="bordagrid" runat="server" style="height: 220px;">
                <asp:GridView ID="GridGruposDeContas" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="GridGruposDeContas_SelectedIndexChanged"
                    Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" ItemStyle-HorizontalAlign="Center" SelectText=" > "
                            ShowSelectButton="True">
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:CommandField>
                        <asp:BoundField DataField="Conta_Id" HeaderText="Grupo">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Titulo" HeaderText="Titulo da Conta">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="bordagrid" runat="server" style="height: 220px;">
                <asp:GridView ID="GridContasAnaliticas" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" ItemStyle-HorizontalAlign="Center" SelectText=" &gt; "
                            ShowSelectButton="True" />
                        <asp:BoundField DataField="Conta_Id" HeaderText="Conta">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Titulo" HeaderText="Titulo">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="row" runat="server">
                <div class="coltxt" style="float: right; padding: 10px;">
                    <asp:Button ID="btnFechar" OnClick="btnFechar_Click" runat="server" CssClass="btn"
                        UseSubmitBehavior="False" Text="Fechar" Style="margin: 0px;" Visible="false">
                    </asp:Button>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
