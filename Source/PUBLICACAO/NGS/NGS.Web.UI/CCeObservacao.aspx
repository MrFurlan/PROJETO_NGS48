<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CCeObservacao.aspx.vb" Inherits="NGS.Web.UI.CCeObservacao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCCeObservacao" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCCeObservacao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                CCe Observações (NFe)
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server" Text="Atualizar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkReimpressao" Text="Reimpressão" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="633px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="633px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="593px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=" > " CssClass="btn" UseSubmitBehavior="False"
                        OnClick="btnCliente_Click" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    UF:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtUf" runat="server" Width="100px" MaxLength="2" />
                </div>
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="100px" />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="100px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtES" runat="server" Width="100px" MaxLength="1" />
                </div>
                <div class="collbl">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNota" Style="text-align: right;" runat="server" class="txtNumerico"
                        Width="100px" />
                </div>
                <div class="collbl" style="margin-left: 21px;">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerie" runat="server" Width="100px" MaxLength="3" />
                </div>
            </div>
            <div class="row">
                <div class="subtitulodiv">
                    Observação:
                </div>
            </div>
            <div class="painelleft" style="width: 49.5%;">
                <div class="row">
                    <div class="collbl">
                        Dados Adicionais:
                    </div>
                </div>
                <div class="row">
                    <div class="coltxt" style="width: 100%;">
                        <asp:TextBox ID="txtDadosAdicionais" runat="server" Width="98%" Height="355px" TextMode="MultiLine" />
                    </div>
                </div>
            </div>
            <div class="painelright" style="width: 49.5%;">
                <div class="row">
                    <div class="collbl">
                        Observações Ficais:
                    </div>
                </div>
                <div class="row">
                    <div class="coltxt" style="width: 100%;">
                        <asp:TextBox ID="txtObservacoesFiscais" runat="server" Width="98%" Height="355px"
                            TextMode="MultiLine" />
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
</asp:Content>
