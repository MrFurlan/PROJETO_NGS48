<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ConsistenciaDeProducao.aspx.vb" Inherits="NGS.Web.UI.ConsistenciaDeProducao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngConsistenciaDeProducao" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlConsistenciaDeProducao" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Consistência de Produção
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
                            </ul>
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
                    <asp:DropDownList ID="ddlUnidade" ToolTip="Unidade de negócio empresarial." runat="server" Width="600px" AutoPostBack="true"
                        OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
            </div>
            <div class="row">
                <div class="collbl">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdAtivo" runat="server" Text="Apenas Produtos Ativos" Checked="True" GroupName="Situacao" data-ToolTip="default" ToolTip="Listar apenas os Produtos com Situação Normal." />
                    <asp:RadioButton ID="rdInativo" runat="server" Text="Apenas Produto(s) Inativo(s)/Bloqueado(s)" GroupName="Situacao" data-ToolTip="default" ToolTip="Listar apenas os Produtos com a Situação diferente da Normal." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Estoque:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkFisico" runat="server" Text="Fisico" data-ToolTip="default" ToolTip="Selecionar estoque físico ou fiscal." />
                </div>
                <div class="coltxt" style="width: 82px;">
                    <asp:CheckBox ID="chkFiscal" runat="server" Text="Fiscal" data-ToolTip="default" ToolTip="Selecionar estoque físico ou fiscal." />
                </div>
                <div class="collbl" style="width: 139px;">
                    <asp:CheckBox ID="CkCDeposito" runat="server" Text="Consolidar Depósito:" data-ToolTip="default" ToolTip="Consolidar Depósito." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período Inicial:
                </div>
                <div class="coltxt" style="width: 135px;">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" data-ToolTip="default" ToolTip="Período inicial da consulta." />
                </div>
                <div class="collbl" style="width: 140px;">
                    Período Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" data-ToolTip="default" ToolTip="Período final da consulta." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
