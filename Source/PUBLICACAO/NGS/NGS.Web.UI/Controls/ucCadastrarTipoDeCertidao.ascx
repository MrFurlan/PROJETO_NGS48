<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucCadastrarTipoDeCertidao.ascx.vb" Inherits="NGS.Web.UI.ucCadastrarTipoDeCertidao" %>

<div id="divCadastrarTipoDeCertidao" class="uc" title="Cadastrar Tipo De Certidão ou Licença" style="display: none">
    <asp:UpdatePanel ID="updpnlCadastrarTipoDeCertidao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="row">
                <div class="collbluc">
                    Descricao:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" runat="server" Width="500px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:Button ID="btnConfirmar" runat="server" ClientIDMode="Static" Text="Confirmar"
                        CssClass="btn" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnFechar" runat="server" ClientIDMode="Static" Text="Fechar" CssClass="btn" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
