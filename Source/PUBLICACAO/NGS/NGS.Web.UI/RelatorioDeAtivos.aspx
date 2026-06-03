<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeAtivos.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeAtivos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 130px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngRelatorioDeAtivos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioDeAtivos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Relatório de Ativos
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
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="626px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="626px" OnSelectedIndexChanged="DdlEmpresa_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlGrupo" runat="server" Width="626px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data da Depreciação:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="100px" data-ToolTip="default" ToolTip="Data de inicio da depreciação." />
                </div>
                <div class="collbl" style="width: 90px;">
                    <asp:CheckBox ID="chkSeguro" Text="Seguro" runat="server" data-ToolTip="default" ToolTip="Selecionar para listar somente os dados de seguros." />
                </div>
                <div class="collbl" style="width: 90px;">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ckDepreciados" runat="server" AutoPostBack="True" OnCheckedChanged="ckDepreciados_CheckedChanged"
                        Text="Depreciados 100%" data-ToolTip="default" ToolTip="Selecionar para listar os bens baixados ou depreciados." />&nbsp;
                    <asp:CheckBox ID="ckBaixados" runat="server" AutoPostBack="True" OnCheckedChanged="ckBaixados_CheckedChanged"
                        Text="Baixados" data-ToolTip="default" ToolTip="Selecionar para listar os bens baixados ou depreciados." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
