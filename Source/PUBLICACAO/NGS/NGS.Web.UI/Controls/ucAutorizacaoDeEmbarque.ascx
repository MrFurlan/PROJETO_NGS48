<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucAutorizacaoDeEmbarque.ascx.vb"
    Inherits="NGS.Web.UI.ucAutorizacaoDeEmbarque" %>
<script type="text/javascript">
    function pageLoadAutorizacaoDeEmbarque() {
        $("#btnConfirmar", "#divAutorizacaoDeEmbarque").button();
        $("#btnFechar", "#divAutorizacaoDeEmbarque").button();
    }

    $(document).ready(function () {
        pageLoadAutorizacaoDeEmbarque();
    });

    var prmAutorizacaoDeEmbarque = Sys.WebForms.PageRequestManager.getInstance();
    prmAutorizacaoDeEmbarque.add_endRequest(pageLoadAutorizacaoDeEmbarque);
</script>
<style type="text/css">
    .size11 {
        font-size: 11px;
    }

    .mleft {
        margin-left: 4px;
    }
    .collbluc{width:108px;}
</style>
<div id="divAutorizacaoDeEmbarque" class="uc" title="Autorização De Embarque" style="display: none;">
    <asp:UpdatePanel ID="updpnlAutorizacaoDeEmbarque" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdnLinhaProd" runat="server" />
            <asp:HiddenField ID="hdnLinhaLocalEntrega" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovoAut" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimparAut" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <%-- *******************1******************************************--%>
            <div class="painelleft" style="width: 58.7%; margin-left: 4px;">
                <div class="subtitulodiv size11">
                    <asp:Label ID="lblLocalEntrega" runat="server" />
                </div>
                <div class="subtitulodiv size11">
                    <asp:Label ID="lblProduto" runat="server" />
                </div>
                <div class="row">
                    <div class="collbluc">
                        Movimento:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMovimentoItemAut" runat="server" CssClass="calendario" Width="77px" />
                    </div>
                    <div class="collbluc">
                        Lançamento:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="cmbTipoLancamentoAut" runat="server" Width="110px" />
                    </div>
                    <div class="collbluc">
                        Quantidade:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtQuantidadeAut" runat="server" CssClass="txtDecimal4" Width="90px" />
                    </div>
                </div>
                <div class="bordagrid" style="height: 188px;">
                    <asp:GridView ID="grd" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:BoundField DataField="NrLancamento" HeaderText="Nr Lançamento" ReadOnly="True">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>
                            <asp:BoundField DataField="TipoDeLancamento" HeaderText="Tipo De Lançamento" ReadOnly="True">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Movimento" HeaderText="Movimento" ReadOnly="True" DataFormatString="{0:dd/MM/yyyy}">
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>
                            <asp:BoundField DataField="UsuarioInclusao" HeaderText="Usuário" ReadOnly="True">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" ReadOnly="True">
                                <HeaderStyle HorizontalAlign="Right" Width="30px" />
                                <ItemStyle HorizontalAlign="Right" Width="30px" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <div class="painelright mleft" style="width: 20%;">
                <div class="subtitulodiv">
                    Resumo Geral do Pedido
                </div>
                <div class="row">
                    <div class="collbluc">
                        Normal:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtNormal" runat="server" BackColor="#FFFFCC" ReadOnly="True" Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Complemento:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtComplemento" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Estorno:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtEstorno" runat="server" BackColor="#FFFFCC" ReadOnly="True" Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Qtde Autorizado:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtQtdeAutorizado" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Embarcado:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtEmbarcado" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Qtde Devolvido:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtQtdeDevolvido" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Max. Autorizar:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMaxAutorizar" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Máx Devolução:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMaxEstornar" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
            </div>
            <div class="painelright mleft" style="width: 20%;">
                <div class="subtitulodiv">
                    Resumo do Local de Entrega
                </div>
                <div class="row">
                    <div class="collbluc">
                        Normal:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtNormalLocal" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Complemento:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtComplementoLocal" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Estorno:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtEstornoLocal" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Qtde Autorizado:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtQtdeAutorizadoLocal" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Embarcado:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtEmbarcadoLocal" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Qtde Devolvido:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtQtdeDevolvidaLocal" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
                <div class="row">
                </div>
                <div class="row">
                    <div class="collbluc">
                        Máx Devolução:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMaxExtornarLocal" runat="server" BackColor="#FFFFCC" ReadOnly="True"
                            Style="text-align: right;" />
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
