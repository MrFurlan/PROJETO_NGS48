<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaObservacoesEmbarque.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaObservacoesEmbarque" %>
<script type="text/javascript">
    function pageLoadConsultaObservacoesEmbarque() {
        $("#btnFechar", "#divConsultaObservacoesEmbarque").button();
    }

    $(document).ready(function () {
        pageLoadConsultaObservacoesEmbarque();
    });

    var prmConsultaObservacoesEmbarque = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaObservacoesEmbarque.add_endRequest(pageLoadConsultaObservacoesEmbarque);
</script>
<div id="divConsultaObservacoesEmbarque" class="uc" title="Consulta de Observações - Observações do Embarque"
    style="display: none;">
    <asp:UpdatePanel ID="updConsultaObservacoesEmbarque" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdLocalDeEntrega" runat="server" />
            <asp:Panel ID="pnlConsultaObservacoes" runat="server" Height="286px" Width="100%">
                <div class="menu_acoes">
                    <div class="acoes">
                        <ul>
                            <li runat="server">
                                <asp:LinkButton ID="lnkConsultarObservacao" Text="Consultar" runat="server" CssClass="iconConsultar" />
                            </li>
                            <li runat="server">
                                <asp:LinkButton ID="lnkSair" Text="Sair" runat="server" CssClass="iconSair" />
                            </li>

                        </ul>
                    </div>
                </div>
                <div class="row">
                    <div class="coltxt" style="width: 99%">
                        <asp:TextBox ID="txtObservacaoEmbarque" runat="server" Height="200px" Width="100%"
                            TextMode="MultiLine" />
                    </div>
                </div>
            </asp:Panel>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
