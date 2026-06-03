<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucLiberacao.ascx.vb" Inherits="NGS.Web.UI.ucLiberacao" %>
<div id="divLiberacao" class="uc" title="Liberação de Senha" style="display: none;">
    <asp:UpdatePanel ID="updpnlLiberacao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkGerar" Text="Gerar" runat="server" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkLiberar" Text="Liberar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <fieldset class="borda_pontilhado">
                    <legend class="fonte6"><strong><em>INFORME ESSE NÚMERO PARA O SUPORTE</em></strong></legend>
                    <asp:Label ID="txtNumero" ClientIDMode="Static" runat="server" ForeColor="Red" Font-Bold="true" />
                </fieldset>
            </div>
            <div class="row">
                <fieldset class="borda_pontilhado">
                    <legend class="fonte6"><strong><em>DIGITE O NÚMERO INFORMADO PELO SUPORTE</em></strong></legend>
                    <div class="coltxt">
                        <asp:TextBox ID="txtNumeroLiberacao" CssClass="txtNumerico" ClientIDMode="Static" runat="server" Width="200px" />
                    </div>
                </fieldset>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>