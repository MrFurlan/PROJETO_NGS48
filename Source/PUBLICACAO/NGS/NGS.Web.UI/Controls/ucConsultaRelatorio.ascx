<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaRelatorio.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaRelatorio" %>
<div id="divConsultaRelatorio" class="uc" title="Escolha o tipo de relatório" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaRelatorio" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <fieldset>
                <legend>Escolha os parâmetros necessários</legend>
                <div class="row">
                    <div class="collbluc">
                        Via de Transporte:
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rdoRodoviario" runat="server" AutoPostBack="true" Checked="true"
                            GroupName="ViaTransporte" Text="Rodoviário" />
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rdoFerroviario" runat="server" AutoPostBack="true" GroupName="ViaTransporte"
                            Text="Ferroviário" />
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rdoRodovXFerrov" runat="server" AutoPostBack="true" GroupName="ViaTransporte"
                            Text="Rodoviário X Ferroviário" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Composição:
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rdoPorNota" runat="server" AutoPostBack="true" GroupName="Composicao"
                            Text="Por Nota" Checked="true" Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rdoPorCtrc" runat="server" AutoPostBack="true" GroupName="Composicao"
                            Text="Por Ctrc" Enabled="false" />
                    </div>
                </div>
            </fieldset>
            <div class="row">
                <div class="painelright">
                    <asp:Button ID="btnGerar" runat="server" CssClass="botao" Text="Gerar" />
                    <asp:Button ID="btnCancelar" runat="server" CssClass="botao" Text="Cancelar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
