<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaTitulo.ascx.vb" Inherits="NGS.Web.UI.ucConsultaTitulo" %>

<script type="text/javascript">
    function pageLoadConsultaRegistro() {

        function consulta() {
            var btn = document.getElementById('<%=lnkConsultar.ClientID%>');
            btn.click();
        }

        $("#MainContent_ucConsultaTitulo_txtRegistro", "#divConsultaTitulo").keydown(function (event) {
            if (event.which == 13) {
                event.preventDefault();
                consulta();
            }
        });

        $("#MainContent_ucConsultaTitulo_txtRegistro", "#divConsultaTitulo").focus();
    }

    $(document).ready(function () {
        pageLoadConsultaRegistro();
    });

    var prmConsultaRegistro = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaRegistro.add_endRequest(pageLoadConsultaRegistro);
</script>

<div id="divConsultaTitulo" class="uc" title="Consultar Título" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaTitulo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Carregar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Registro:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtRegistro" runat="server" CssClass="txtNumerico9" Style="text-align: right;"
                        Width="150px" Font-Size="10pt" Font-Bold="True" data-ToolTip="default" ToolTip="Número do lançamento." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>

