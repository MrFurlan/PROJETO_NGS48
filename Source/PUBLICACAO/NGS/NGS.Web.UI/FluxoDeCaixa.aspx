<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="FluxoDeCaixa.aspx.vb" Inherits="NGS.Web.UI.FluxoDeCaixa" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 120px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManagerFlxCx" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="UpdatePanelFlxCx" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Fluxo de Caixa
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
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
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" Style="width: 590px;" runat="server" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" Style="width: 590px;" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período de:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData1" CssClass="calendario" runat="server" Width="100px" data-ToolTip="default"
                        ToolTip="Data inicial à data final da consulta." />
                    a
                    <asp:TextBox ID="txtData2" runat="server" CssClass="calendario" Width="100px" data-ToolTip="default"
                        ToolTip="Data inicial à data final da consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedidos Emprestimo:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkEmprestimo" Text="SIM" runat="server" Checked="true" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
