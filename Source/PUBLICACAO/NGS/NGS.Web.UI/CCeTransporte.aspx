<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CCeTransporte.aspx.vb" Inherits="NGS.Web.UI.CCeTransporte" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 980px !important;
        }

        .width {
            width: 130px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCCeTransporte" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCCeTransporte" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                CCe Transportador e Placa(s) (NFe)
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkReimpressao" Text="Reimpressão" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
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
                    <asp:Button ID="btnCliente" CssClass="btn" runat="server" Text=" > " UseSubmitBehavior="False"
                        OnClick="btnCliente_Click" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="79px"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="79px"
                        data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtES" runat="server" Width="100px" MaxLength="1" data-ToolTip="default"
                        ToolTip="" />
                </div>
                <div class="collbl">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNota" Style="text-align: right;" runat="server" class="txtNumerico"
                        Width="100px" data-ToolTip="default" ToolTip="Número da Nota Fiscal." />
                </div>
                <div class="collbl">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSerie" runat="server" Width="100px" MaxLength="3" data-ToolTip="default"
                        ToolTip="Série da Nota Fiscal." />
                </div>
            </div>
            <div class="subtitulodiv">
                Dados do transportador
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkTransportador" runat="server" Enabled="False" Visible="False" />Transportador:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoTransportador" runat="server" />
                    <asp:TextBox ID="txtTransportador" runat="server" Width="591px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnTransportador" CssClass="btn" runat="server" Text=">" UseSubmitBehavior="False"
                        Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkVeiculo" runat="server" Enabled="False" />Veículo:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoPlaca" runat="server" />
                    <asp:TextBox ID="txtVeiculo" runat="server" MaxLength="20" Enabled="False" Width="80px" />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgPlaca" runat="server" data-ToolTip="default" ToolTip="Procura Placa do Veículo"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ImageAlign="AbsMiddle" />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="chkEstadoVeiculo" runat="server" Visible="False" Enabled="False" />Estado:
                </div>
                <div class="coltxt width">
                    <asp:Label ID="lblEstadoVeiculo" runat="server" Text="" />&nbsp;
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgEstadoVeiculo" runat="server" data-ToolTip="default" ToolTip="Procura Estado do Veículo"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ImageAlign="AbsMiddle" />
                </div>
                <div class="collbl">
                    Cidade:
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblCidadeVeiculo" runat="server" Text="" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkReboque1" runat="server" Enabled="False" />Reboque 1
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtReboque1" runat="server" MaxLength="20" Enabled="False" Width="80px" />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="chkEstadoReboque1" runat="server" Enabled="False" />Estado:
                </div>
                <div class="coltxt width">
                    <asp:Label ID="lblEstadoReboque1" runat="server" Text="" />&nbsp;
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgEstadoReboque1" runat="server" data-ToolTip="default" ToolTip="Procura Estado do Reboque 1"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ImageAlign="AbsMiddle"
                        Visible="False" />
                </div>
                <div class="collbl">
                    Cidade:
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblCidadeReboque1" runat="server" Text="" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkReboque2" runat="server" Enabled="False" />Reboque 2
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtReboque2" runat="server" MaxLength="20" Enabled="False" Width="80px" />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="chkEstadoReboque2" runat="server" Enabled="False" />Estado:
                </div>
                <div class="coltxt width">
                    <asp:Label ID="lblEstadoReboque2" runat="server" Text="" />&nbsp;
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgEstadoReboque2" runat="server" data-ToolTip="default" ToolTip="Procura Estado do Reboque 2"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ImageAlign="AbsMiddle"
                        Visible="False" />
                </div>
                <div class="collbl">
                    Cidade:
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblCidadeReboque2" runat="server" Text="" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkReboque3" runat="server" Enabled="False" />Reboque 3
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtReboque3" runat="server" MaxLength="20" Enabled="False" Width="80px" />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="chkEstadoReboque3" runat="server" Enabled="False" />Estado:
                </div>
                <div class="coltxt width">
                    <asp:Label ID="lblEstadoReboque3" runat="server" Text="" />&nbsp;
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgEstadoReboque3" runat="server" data-ToolTip="default" ToolTip="Procura Estado do Reboque 3"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ImageAlign="AbsMiddle"
                        Visible="False" />
                </div>
                <div class="collbl">
                    Cidade:
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblCidadeReboque3" runat="server" Text="" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPlacas ID="ucConsultaPlacas" runat="server" />
    <uc:ConsultaEstados ID="ucConsultaEstados" runat="server" />
    <uc:ConsultaCodMunicipios ID="ucConsultaCodMunicipios" runat="server" />
</asp:Content>
