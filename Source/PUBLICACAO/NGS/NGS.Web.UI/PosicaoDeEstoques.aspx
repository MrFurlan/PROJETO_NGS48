<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PosicaoDeEstoques.aspx.vb" Inherits="NGS.Web.UI.PosicaoDeEstoques" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPosicaoDeEstoques" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPosicaoDeEstoques" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Posição de Estoques
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
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcelDados" runat="server" Text="ExcelDados" />
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
                    <asp:DropDownList ID="ddlUnidade" ToolTip="Unidade de negócio empresarial." runat="server" Width="600px" Enabled="False" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="600px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlDeposito" runat="server" Width="600px" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Estoque:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadFisico" runat="server" Enabled="False" Text=" Fisico " GroupName="Estoque" data-ToolTip="default" ToolTip="Selecionar estoque físico ou fiscal." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadFiscal" runat="server" Enabled="False" Text=" Fiscal " GroupName="Estoque"
                        Checked="True" data-ToolTip="default" ToolTip="Selecionar estoque físico ou fiscal." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" runat="server" CssClass="calendario" data-ToolTip="default" ToolTip="Período da consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="CkCDeposito" runat="server" Text="Depósito:" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkConsNotasComSerieEspecificas" runat="server" Checked="True"
                        Text="Considerar Notas C/ Series 101, 102, 103, 104" data-ToolTip="default" ToolTip="Marcar a opção para gerar o relatório desejado." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkContraPartida" runat="server" Text="Gerar Contra Partida dos Depositos de Terceiros" data-ToolTip="default" ToolTip="Marcar a opção para gerar o relatório desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdAtivo" runat="server" Text="Apenas Produtos Ativos" Checked="True" GroupName="Situacao" data-ToolTip="default" ToolTip="Listar apenas os Produtos com Situação Normal." />
                    <asp:RadioButton ID="rdInativo" runat="server" Text="Apenas Produto(s) Inativo(s)/Bloqueado(s)" GroupName="Situacao" data-ToolTip="default" ToolTip="Listar apenas os Produtos com a Situação diferente da Normal." />
                    <asp:CheckBox ID="chkAlmoxarifado" runat="server" Text="Apenas Almoxarifado" data-ToolTip="default" ToolTip="Marcar a opção para gerar o relatório desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Ordem do Produto:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdAlfabetico" runat="server" Text="Nome" GroupName="OrdemRelatorio"
                        Checked="True" data-ToolTip="default" ToolTip="Ordenado por nome" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdProduto" runat="server" Text="Código" GroupName="OrdemRelatorio" data-ToolTip="default" ToolTip="Ordenado por código" />
                </div>
            </div>
            <div class="row">
                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
