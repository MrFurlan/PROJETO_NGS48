<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ContabilizacaoDeNotas.aspx.vb" Inherits="NGS.Web.UI.ContabilizacaoDeNotas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngContabilizacaoDeNotas" runat="server" EnableScriptGlobalization="True" EnableScriptLocalization="True" AsyncPostBackTimeout="100000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlContabilizacaoDeNotas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Contabilização de Notas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkProcessar" Text="Processar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkContabilizar" Text="Contabilizar Notas" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <%--<div class="row">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RdNota" runat="server" AutoPostBack="True" Checked="True" GroupName="Razao"
                        OnCheckedChanged="RdNota_CheckedChanged" Text="Notas" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdFrete" runat="server" AutoPostBack="True" GroupName="Razao"
                        OnCheckedChanged="rdFrete_CheckedChanged" Visible="false" Text="Transferencia de Frete" />
                </div>
            </div>--%>
            <div class="row">
                <div class="collbl">
                    <asp:Label ID="Label1" runat="server" Text="Unidade" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="600px" AutoPostBack="true"
                        OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:Label ID="Label2" runat="server" Text="Empresa" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" />
                </div>
            </div>
            <div id="divNota" runat="server" class="row">
                <div class="collbl">
                    Un.Contratante:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeContratante" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeContratante_SelectedIndexChanged"
                        Width="600px" />
                </div>
            </div>
            <div id="divNota2" runat="server" class="row">
                <div class="collbl">
                    Contratante:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresaContratante" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox Style="width: 80px;" ID="TxtdataInicial" runat="server" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtdataFinal" Style="width: 80px;" runat="server" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" data-ToolTip="default" ToolTip="Número do pedido a ser consultado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Informações:
                </div>
                <div class="coltxt">
                    <asp:CheckBoxList ID="ChkResultado" runat="server" Enabled="False" Font-Bold="True"
                        data-ToolTip="default" ToolTip="Informações referente a execução do processo de contabilização." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ContabilizarNotas ID="ucContabilizarNotas" runat="server" />
</asp:Content>
