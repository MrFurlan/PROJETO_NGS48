<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="UcInputDate.ascx.vb"
    Inherits="NGS.Web.UI.UcInputDate" %>
<div id="divInputDate" class="uc" title="Informe as Datas" style="display: none;">
    <asp:UpdatePanel ID="updpnlInputDate" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Informe Ano Inicial e Ano Final para inclusao de datas
            </div>
            <div class="row">
                <div class="collbl">
                    Ano Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAno1" runat="server" CssClass="txtNumerico" Style="width: 90px;" />
                </div>
                <div class="collbl">
                    Ano Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAno2" runat="server" CssClass="txtNumerico" Style="width: 90px;" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="margin-right: 4px;">
                    <asp:Button ID="btnFechar" runat="server" ClientIDMode="Static" Text="Fechar" CssClass="botao" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConfirmar" runat="server" ClientIDMode="Static" Text="Confirmar"
                        CssClass="botao" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
