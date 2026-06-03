<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucSupervisor.ascx.vb"
    Inherits="NGS.Web.UI.ucSupervisor" %>
<script type="text/javascript">
    function pageLoadSupervisor() {
        $("#MainContent_ucSupervisor_txtUsuario", "#divSupervisor").focus();


        $("#MainContent_ucSupervisor_txtSenha", "#divSupervisor").keydown(function (event) {
            if (event.which == 13) {
                event.preventDefault();
                $("#MainContent_ucSupervisor_btnEntrar", "#divSupervisor").click();
            }
        });

    }
    $(document).ready(function () {
        pageLoadSupervisor();
    });

    var prmSupervisor = Sys.WebForms.PageRequestManager.getInstance();
    prmSupervisor.add_endRequest(pageLoadSupervisor);




</script>
<div id="divSupervisor" class="uc" title="Adicione as Credenciais de Supervisor"
    style="display: none;">
    <%--style="display: none;"--%>
    <asp:UpdatePanel ID="updpnlSupervisor" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div style="width: 335px;">
                <div style="background: none repeat scroll 0 0 #F5F5F5; border: 1px solid #E5E5E5;
                    margin: 12px 0 0; padding: 20px 25px 15px;">
                    <h2 style="font-size: 15pt; height: 16px; line-height: 16px; margin: 0 0 1.2em; position: relative;">
                        Login <strong style="background: url(noimagens/ngspqn.png) no-repeat scroll 0 0 transparent;
                            display: inline-block; height: 19px; position: absolute; right: 0; top: 1px;
                            width: 52px;"></strong>
                    </h2>
                    <div style="text-align: center;">
                        <asp:Label ID="lblErro" runat="server" ForeColor="Red" Font-Bold="true" Visible="false" />
                    </div>
                    <br />
                    <label style="display: block; margin: 0 0 1.5em;">
                        <strong style="-moz-user-select: none; display: block; font-weight: bold; margin: 0 0 0.5em;
                            font-size: 12pt">Nome de usuário</strong>
                        <asp:TextBox ID="txtUsuario" runat="server" Style="font-size: 12pt; height: 32px;
                            width: 100%;" />
                    </label>
                    <label style="display: block; margin: 0 0 1.5em;">
                        <strong style="-moz-user-select: none; display: block; font-weight: bold; margin: 0 0 0.5em;
                            font-size: 12pt">Senha</strong>
                        <asp:TextBox ID="txtSenha" runat="server" Style="font-size: 12pt; height: 32px; width: 100%;"
                            TextMode="Password" />
                    </label>
                    <label style="display: block; margin: 0 0 1.5em;">
                        <asp:DropDownList ID="cmbBancoServidor" runat="server" Visible="False" Style="font-size: 8pt;
                            height: 32px; width: 289px;" />
                    </label>
                    <asp:Button ID="btnEntrar" runat="server" Text="Entrar" OnClick="btnEntrar_Click"
                        CssClass="botaologin" Height="38px" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btnCancelar" runat="server" Text="Cancelar" CssClass="botaologin"
                        Height="38px" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
