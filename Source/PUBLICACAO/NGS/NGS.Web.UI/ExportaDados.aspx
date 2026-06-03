<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ExportaDados.aspx.vb" Inherits="NGS.Web.UI.ExportaDados" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmExportaDados" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlExportaDados" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Exportações de Arquivos
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconConfirmar" ID="lnkConfirmar" Text="Confirmar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconSair" ID="lnkCancelar" Text="Cancelar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConfirmar" ID="lnkFinalizar" Text="Finalizar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 155px;">
                    <input id="checkOrcamentos" type="checkbox" />Solicitações de Compras:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtOrcamentos" runat="server" Width="478px" ReadOnly="true" TabIndex="1"
                        data-ToolTip="default" ToolTip="Selecionar para exportar as solicitações de compras." />
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 155px;">
                    <input id="checkCotacoes" type="checkbox" tabindex="2" />Cotações de Compras:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtCotacoes" runat="server" Width="480px" ReadOnly="true" TabIndex="3"
                        data-ToolTip="default" ToolTip="Selecionar para exportar as cotações de compras." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
