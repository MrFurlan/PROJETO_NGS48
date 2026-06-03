<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" 
    CodeBehind="RelatorioDeComissoesPedido.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeComissoesPedido" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngMonitorDeXML" runat="server" EnableScriptGlobalization="True" EnableScriptLocalization="True" AsyncPostBackTimeout="5000">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlMonitorDeXML" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Relatório De Comissões por Pedido
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkPosição" Text="Posição" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkExcel" Text="Excel" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkResumoExcel" Text="Resumo Excel" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajudar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="618px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="collbl" style="width: 150px;">
                    <asp:CheckBox ID="chkConsolidarEmpresa" runat="server" Text="Consolidar Empresa"
                            data-ToolTip="default" ToolTip="Consolidar o cnpj da empresa." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Representante:
                </div>
                <div class="collbl" style="width: 150px;">
                    <asp:CheckBox ID="chkTodosRepresentantes" runat="server" Text="Todos Representantes"
                        data-ToolTip="default" ToolTip="Todos os representantes serão selecionados." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlRepresentante" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" Width="80px" CssClass="calendario" 
                        data-ToolTip="default" ToolTip="Informar o período inicial da consulta." />
                </div>
                <div class="collbl">
                    à
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" Width="80px" CssClass="calendario" 
                        data-ToolTip="default" ToolTip="Informar o período final da consulta." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
