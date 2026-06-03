<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucDestinoContabil.ascx.vb"
    Inherits="NGS.Web.UI.ucDestinoContabil" %>
<script type="text/javascript">
    function pageLoadDestinoContabil() {
        $("#btnFechar", "#divDestinoContabil").button();
    }

    $(document).ready(function () {
        pageLoadDestinoContabil();
    });

    var prmDestinoContabil = Sys.WebForms.PageRequestManager.getInstance();
    prmDestinoContabil.add_endRequest(pageLoadDestinoContabil);
</script>
<div id="divDestinoContabil" class="uc" title="Consulta de Destino Contábil" style="display: none;">
    <center>
        <asp:UpdatePanel ID="updpnlDestinoContabil" runat="server">
            <ContentTemplate>
                <asp:HiddenField ID="HID" runat="server" />
                <table width="100%" border="0" cellpadding="0">
                    <tr>
                        <td style="white-space: nowrap; width: 150px;">
                            <div class="headerGray">
                                <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Cliente:</span>
                                <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                            </div>
                        </td>
                        <td>
                            <asp:HiddenField ID="txtCodigoClienteDestino" runat="server" />
                            <table cellpadding="0" width="100%">
                                <tr>
                                    <td>
                                        <asp:TextBox ID="txtClienteDestino" runat="server" ReadOnly="True" Width="98%"/>
                                    </td>
                                    <td>
                                        <asp:Button ID="btnCliente" runat="server" Text=" > " UseSubmitBehavior="False" OnClick="btnCliente_Click" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="white-space: nowrap; width: 150px;">
                            <div class="headerGray">
                                <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Finalidade Financeira:</span>
                                <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                            </div>
                        </td>
                        <td>
                            <asp:DropDownList ID="ddlCarteiraDestino" runat="server" 
                                 Width="100%">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" align="right" style="padding-top: 10px;">
                            <table cellpadding="0">
                                <tr>
                                    <td>
                                        <asp:Button ID="btnConfirmar" runat="server" CssClass="botao" Text="Confirmar" UseSubmitBehavior="False"
                                            OnClick="btnConfirmar_Click" Width="85px" />
                                    </td>
                                    <td>
                                        <asp:Button ID="btnSair" runat="server" CssClass="botao" Text="Fechar" UseSubmitBehavior="False"
                                            OnClick="btnSair_Click" Width="85px" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </ContentTemplate>
        </asp:UpdatePanel>
    </center>
</div>
