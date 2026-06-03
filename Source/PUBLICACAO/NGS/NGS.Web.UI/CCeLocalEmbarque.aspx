<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="CCeLocalEmbarque.aspx.vb" Inherits="NGS.Web.UI.CCeLocalEmbarque" %>
<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .painelleft
        {
            width: 150px;
            margin-right: 4px;
        }
    </style>
</asp:Content><asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
        <asp:ScriptManager ID="scmngCCeLocalEmbarque" runat="server" AsyncPostBackTimeout="50000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCCeLocalEmbarque" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                CCe Local de Embarque (NFe)
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
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="630px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="630px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="591px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=" > " UseSubmitBehavior="False" CssClass="btn"
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
                    <asp:TextBox ID="txtNota" runat="server" Style="text-align: right;" class="txtNumerico"
                        Width="100px" />
                </div>
                <div class="collbl" style="margin-left: 21px;">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerie" runat="server" Width="100px" MaxLength="3" />
                </div>
            </div>
            <hr />
            <div class="row">
                <div class="collbl">
                    Atual:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoLocEmbAtual" runat="server" />
                    <asp:TextBox ID="txtLocEmbAtual" runat="server" Width="591px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:TextBox ID="txtUFLocEmbAtual" runat="server" Width="50px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEnderecoLocEmbAtual" runat="server" Width="652px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Novo:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoLocEmbNovo" runat="server" />
                    <asp:TextBox ID="txtLocEmbNovo" runat="server" Width="591px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnLocEmbNovo" runat="server" Text=" > " UseSubmitBehavior="False" CssClass="btn" Enabled="False" OnClick="btnLocEmbNovo_Click" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:TextBox ID="txtUFLocEmbNovo" runat="server" Width="50px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEnderecoLocEmbNovo" runat="server" Width="652px" Enabled="False" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />

</asp:Content>
