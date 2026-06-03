<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucOrigemDestino.ascx.vb"
    Inherits="NGS.Web.UI.ucOrigemDestino" %>
<div id="divOrigemDestinoRoteiro" class="uc" title="Cadastro de Roteiros" style="display: none;">
    <asp:UpdatePanel ID="updpnlOrigemDestino" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConfirmar" runat="server">
                            <asp:LinkButton ID="lnkConfirmar" Text="Confirmar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconSair" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Origem:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="hdfCodigoOrigem" runat="server" />
                    <asp:TextBox ID="txtNomeOrigem" runat="server" Width="600px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnOrigem" runat="server" Text=">" CssClass="btn" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Destino:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="hdfCodigoDestino" runat="server" />
                    <asp:TextBox ID="txtNomeDestino" runat="server" Width="600px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnDestino" runat="server" Text=">" CssClass="btn" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Via de Transporte:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlViaTransporte" runat="server" Style="width: 280px;" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
