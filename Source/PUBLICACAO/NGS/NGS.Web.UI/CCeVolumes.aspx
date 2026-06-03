<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CCeVolumes.aspx.vb" Inherits="NGS.Web.UI.CCeVolumes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .txtDecimal4
        {
            text-align: right;
        }
        .painelleft
        {
            width: 150px;
            margin-right: 4px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCCeObservacao" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCCeObservacao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                CCe Volumes Transportados (NFe)
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
            <div class="painelleft" style="margin-left: 15px;">
                <div class="subtitulodiv">
                    <asp:CheckBox ID="chkVolume" runat="server" AutoPostBack="True" />QUANTIDADE
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtVolumes" runat="server" CssClass="txtDecimal4" ForeColor="Blue"
                        Width="140px" Enabled="False" />
                </div>
            </div>
            <div class="painelleft">
                <div class="subtitulodiv">
                    <asp:CheckBox ID="chkEspecie" runat="server" AutoPostBack="True" />ESPÉCIE
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEspecie" Width="140px" runat="server" ForeColor="Blue" Enabled="False" />
                </div>
            </div>
            <div class="painelleft">
                <div class="subtitulodiv">
                    <asp:CheckBox ID="chkMarca" runat="server" AutoPostBack="True" />MARCA
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtMarca" runat="server" Width="140px" ForeColor="Blue" Enabled="False" />
                </div>
            </div>
            <div class="painelleft">
                <div class="subtitulodiv">
                    <asp:CheckBox ID="chkNumeracao" runat="server" AutoPostBack="True" />NUMERAÇÃO
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumeracao" runat="server" Width="140px" ForeColor="Blue" Enabled="False" />
                </div>
            </div>
            <div class="painelleft">
                <div class="subtitulodiv">
                    <asp:CheckBox ID="chkBruto" runat="server" AutoPostBack="True" />PESO BRUTO
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPesoBruto" CssClass="txtDecimal3" runat="server" ForeColor="Blue"
                        Enabled="False" Width="140px" />
                </div>
            </div>
            <div class="painelleft">
                <div class="subtitulodiv">
                    <asp:CheckBox ID="chkLiquido" runat="server" AutoPostBack="True" />PESO LÍQUIDO
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPesoLiquido" CssClass="txtDecimal3" runat="server" ForeColor="Blue"
                        Width="140px" Enabled="False" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
</asp:Content>
