<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeLaudosPorPlacas.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeLaudosPorPlacas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioDeLaudosPorPlacas" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioDeLaudosPorPlacas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório de Laudos por Placa
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="590px" OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="590px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" Enabled="false" runat="server" Width="550px" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=">" OnClick="btnCliente_Click" UseSubmitBehavior="False"
                        CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Transportadora:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTransportador" Enabled="false" runat="server" Width="550px" />
                    <asp:HiddenField ID="txtCodigoTransportador" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnTransportador" runat="server" Text=">" OnClick="btnTransportador_Click"
                        CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Nome da pessoa/empresa responsável pelo transporte da mercadoria." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Placa:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtPlaca" Width="90px" runat="server" data-ToolTip="default" ToolTip="Identificação do veículo de transporte." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkTodas" runat="server" Text="Buscar em todas..." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" Width="85px" CssClass="calendario" runat="server"
                        data-ToolTip="default" ToolTip="Informar o período inicial de consulta." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" Width="85px" CssClass="calendario" runat="server"
                        data-ToolTip="default" ToolTip="Informar o período final de consulta." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
