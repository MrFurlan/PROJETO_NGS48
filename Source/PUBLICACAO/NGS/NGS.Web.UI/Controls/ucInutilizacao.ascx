<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucInutilizacao.ascx.vb"
    Inherits="NGS.Web.UI.ucInutilizacao" %>
<script type="text/javascript">

    function pageLoadInutilizacao() {
        $('#MainContent_ucInutilizacao_txtJustificativa').attr('maxLength', 255);
    }

    $(document).ready(function () {
        pageLoadInutilizacao();
    });

    var prmInutilizacao = Sys.WebForms.PageRequestManager.getInstance();
    prmInutilizacao.add_endRequest(pageLoadInutilizacao);
</script>
<div id="divInutilizacao" class="uc" title="Inutilização de sequência da nota fiscal"
    style="display: none;">
    <asp:UpdatePanel ID="updpnlInutilizacao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdfTipoDeDocumento" runat="server" />
            <asp:HiddenField ID="hdLiberar" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkLiberar" Text="Liberar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconFechar" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEmpresa" runat="server" Width="545px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnEmpresa" runat="server" Text=">" CssClass="btn" OnClick="btnEmpresa_Click" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Nota Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNotaInicial" runat="server" CssClass="txtNumerico" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Nota Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNotaFinal" runat="server" CssClass="txtNumerico" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerie" runat="server" MaxLength="3"/>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdEntrada" runat="server" GroupName="ES" Text="Entrada" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdSaida" runat="server" Checked="True" GroupName="ES" Text="Saída" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Justificativa:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtJustificativa" runat="server" Height="45px" TextMode="MultiLine"
                        Width="580px" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
