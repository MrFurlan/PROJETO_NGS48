<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioParticipacaoNaVendaLucroBruto.aspx.vb" Inherits="NGS.Web.UI.RelatorioParticipacaoNaVendaLucroBruto" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 150px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioParticipacaoNaVendaLucroBruto" runat="server"
        EnableScriptGlobalization="True" EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioParticipacaoNaVendaLucroBruto" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório Participação Na Venda / Lucro Bruto
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
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
                    Relatório:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbPorCliente" runat="server" Checked="True" GroupName="filtro"
                        Text="Por Cliente" data-ToolTip="default" ToolTip="Selecionar de acordo com o relatório desejado." />
                    <asp:RadioButton ID="rbPorProduto" runat="server" GroupName="filtro" Text="Por Produto" data-ToolTip="default" ToolTip="Selecionar de acordo com o relatório desejado." />
                    <asp:RadioButton ID="rbPorRepresentante" runat="server" GroupName="filtro" Text="Por Representante" data-ToolTip="default" ToolTip="Selecionar de acordo com o relatório desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEmpresa" runat="server" Width="618px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente/Representante:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliRepre" runat="server" />
                    <asp:TextBox ID="txtClienteRepresentante" runat="server" Enabled="False" Width="578px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaCliRepre" runat="server" CausesValidation="False" OnClick="cmdBuscaCliOrigem_Click"
                        CssClass="btn" Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o representante/cliente para consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Consolidar:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ckConsolidarEmp" runat="server" Text="Empresa" data-ToolTip="default" ToolTip="Marcar a opção desejada para consolidar." />
                    <asp:CheckBox ID="ckConsolidarCli" runat="server" Text="Cliente / Representante" data-ToolTip="default" ToolTip="Marcar a opção desejada para consolidar." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Marca:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMarca" runat="server" Width="301px" />
                </div>
            </div>
            <div class="row" runat="server">
                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="300px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="112px" data-ToolTip="default" ToolTip="Data Inicial a final da consulta." />
                    &nbsp;á &nbsp;
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="112px" data-ToolTip="default" ToolTip="Data Inicial a final da consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Considerar:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ckCPeriodo" runat="server" Text="Considerar Período" data-ToolTip="default" ToolTip="Selecionar para considerar o período." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Ordenar por:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbVenda" runat="server" Checked="True" GroupName="participacao"
                        Text="Participação na Venda" data-ToolTip="default" ToolTip="Selecionar para ordenar por participação na venda ou no lucro." />
                    <asp:RadioButton ID="rbLucro" runat="server" GroupName="participacao" Text="Participação no Lucro"
                        data-ToolTip="default" ToolTip="Selecionar para ordenar por participação na venda ou no lucro." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
