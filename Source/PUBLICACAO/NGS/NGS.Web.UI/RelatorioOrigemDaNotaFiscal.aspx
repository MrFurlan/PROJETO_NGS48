<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioOrigemDaNotaFiscal.aspx.vb" Inherits="NGS.Web.UI.RelatorioOrigemDaNotaFiscal" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmRelacaoDePedidos" runat="server" AsyncPostBackTimeout="10000"
        EnableScriptGlobalization="True" EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelacaoDePedidos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório de Origens da Nota Fiscal
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" runat="server">
                                <span>Relatório</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server">
                                <span>Limpar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server">
                                <span>Ajuda</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" AutoPostBack="true" Width="590px" />
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
                    <asp:TextBox ID="txtCliente" runat="server" Width="550px" Enabled="false" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" CssClass="btn" runat="server" Text=">" OnClick="btnCliente_Click"
                        UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nota:
                </div>
                <div class="coltxt" style="width: 131px;">
                    <asp:TextBox runat="server" ID="txtNota" Width="100px" data-ToolTip="default" ToolTip="Inserir o número da Nota Fiscal." />
                </div>
                <div class="collbl">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" ID="txtSerie" Width="100px" data-ToolTip="default" ToolTip="Inserir a série da Nota Fiscal." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data de:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" CssClass="calendario" ID="txtData1" Width="100px" data-ToolTip="default"
                        ToolTip="Informar o período inicial de consulta." />
                </div>
                <div class="collbl">
                    Até:
                </div>
                <div class="coltxt">
                    <asp:TextBox runat="server" CssClass="calendario" ID="txtData2" Width="100px" data-ToolTip="default"
                        ToolTip="Informar o período final de consulta." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
