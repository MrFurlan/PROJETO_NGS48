<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CustoMedio.aspx.vb" Inherits="NGS.Web.UI.CustoMedio" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCustoMedio" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCustoMedio" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Custo Médio
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton class="iconPdf" ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcelDados" runat="server" Text="Excel Dados" />
                                </li>
                            </ul>
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="600px" OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlDeposito" runat="server" Width="600px" />
                </div>
            </div>
            <div style="margin-top: 3px;">
                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlMes" runat="server" Width="104px" OnSelectedIndexChanged="DdlMes_SelectedIndexChanged"
                        AutoPostBack="True">
                        <asp:ListItem Value="01">Janeiro</asp:ListItem>
                        <asp:ListItem Value="02">Fevereiro</asp:ListItem>
                        <asp:ListItem Value="03">Mar&#231;o</asp:ListItem>
                        <asp:ListItem Value="04">Abril</asp:ListItem>
                        <asp:ListItem Value="05">Maio</asp:ListItem>
                        <asp:ListItem Value="06">Junho</asp:ListItem>
                        <asp:ListItem Value="07">Julho</asp:ListItem>
                        <asp:ListItem Value="08">Agosto</asp:ListItem>
                        <asp:ListItem Value="09">Setembro</asp:ListItem>
                        <asp:ListItem Value="10">Outubro</asp:ListItem>
                        <asp:ListItem Value="11">Novembro</asp:ListItem>
                        <asp:ListItem Value="12">Dezembro</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp; / &nbsp;
                    <asp:DropDownList ID="DdlAno" runat="server" Width="120px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkAuditoria" runat="server" Text="Auditoria De:" AutoPostBack="true"
                        data-ToolTip="default" ToolTip="Selecionar em casos de auditoria de notas, razão ou produção." /></span>
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdNotas" runat="server" Text="Custo - Notas" GroupName="auditoria"
                        Enabled="false" data-ToolTip="default" ToolTip="Selecionar em casos de auditoria de notas, razão ou produção." />
                    <asp:RadioButton ID="rdRazao" runat="server" Text="Custo - Razão" GroupName="auditoria"
                        Enabled="false" data-ToolTip="default" ToolTip="Selecionar em casos de auditoria de notas, razão ou produção." />
                    <asp:RadioButton ID="rdProducao" runat="server" Text="Custo - Produção" GroupName="auditoria"
                        Enabled="false" data-ToolTip="default" ToolTip="Selecionar em casos de auditoria de notas, razão ou produção." />
                </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
