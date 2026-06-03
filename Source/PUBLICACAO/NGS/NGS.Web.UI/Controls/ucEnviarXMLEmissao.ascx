<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucEnviarXMLEmissao.ascx.vb"
    Inherits="NGS.Web.UI.ucEnviarXMLEmissao" %>

<div id="divEnviarXMLEmissao" class="uc" title="Enviar XML para emissão de nota fiscal" style="display: none;">
    <asp:UpdatePanel ID="updpnlEnviarXMLEmissao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton ID="lnkFechar" CssClass="iconSair" runat="server" Text="Fechar" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:RadioButton ID="rdNotaFiscal" runat="server" Text="Notas Fiscais" Checked="false"
                        data-ToolTip="default" ToolTip="Carregar Notas Fiscais." GroupName="GrupTipo"/>
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdNotaGeral" runat="server" Text="Notas Gerais" Checked="false"
                        data-ToolTip="default" ToolTip="Carregar Notas Gerais." GroupName="GrupTipo"/>
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdConhecimentoTransporte" runat="server" Text="Conhecimento Transporte" Checked="false"
                        data-ToolTip="default" ToolTip="Carregar CTE - Conhecimento de Transporte" GroupName="GrupTipo" Enabled="false"/>
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:Button ID="btnEnviar" runat="server" ClientIDMode="Static" Text="Enviar"
                        CssClass="btn" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
