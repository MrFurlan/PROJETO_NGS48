<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucCadastrarGrupo.ascx.vb"
    Inherits="NGS.Web.UI.ucCadastrarGrupo" %>
<div id="divCadastrarGrupo" class="uc" title="Cadastrar Grupo" style="display: none;">
    <asp:UpdatePanel ID="updpnlCadastrarGrupo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="row">
                <div class="collbluc">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtGrupo" runat="server" Width="500px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" TabIndex="2" runat="server" Width="500px" />
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
