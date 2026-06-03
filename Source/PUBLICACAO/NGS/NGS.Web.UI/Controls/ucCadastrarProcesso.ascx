<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucCadastrarProcesso.ascx.vb"
    Inherits="NGS.Web.UI.ucCadastrarProcesso" %>
<div id="divCadastrarProcesso" class="uc" title="Cadastrar Processo" style="display: none;">
    <asp:UpdatePanel ID="updpnlCadastrarProcesso" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="row">
                <div class="collbluc">
                    Processo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtProcesso" Style="text-transform: uppercase;" runat="server" Width="500px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" Style="text-transform: uppercase;" TabIndex="2" runat="server"
                        Width="500px" />
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
